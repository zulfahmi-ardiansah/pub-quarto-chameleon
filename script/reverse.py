"""Extract a .docx (or .md) input into clean Markdown pages (qlon ingest path).

Usage:
    qlon <input.docx|input.md> [--use-llm] [--layout flat|fuma]

Extracts the .docx to Markdown with `quarto pandoc`, splits it at headings, then
(with --use-llm) cleans each section via an LLM or, by default, passes the raw text
through untouched. Writes one page file per H1 plus a media folder. Prompts live in
config/prompt.py; the LLM is configured via LLM_API_KEY / LLM_BASE_URL / LLM_MODEL.
"""
from __future__ import annotations
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import uuid
from dataclasses import dataclass
from datetime import date
from pathlib import Path

import httpx
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn
from rich.text import Text

from utility import (
    DEFAULT_MODEL,
    LLMError,
    api_key,
    chat_completion,
    default_model,
    load_dotenv,
    slugify,
)

# Prompt constants live in config/ (not on sys.path by default); make it importable.
ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR / "config"))
from prompt import (  # noqa: E402  (path is set up just above)
    CLEAN_SYSTEM_PROMPT,
    STRUCTURE_PROMPT,
    DESCRIBE_PROMPT,
    SEGMENT_PROMPT,
)

# Shared console so diagnostic lines render cleanly above the live progress bar.
_console = Console()


def _info(msg: str) -> None:
    """Print a diagnostic line through the shared console (markup/highlight off so
    arbitrary content -- reprs, brackets, lists -- is shown verbatim, never parsed)."""
    _console.print(msg, markup=False, highlight=False)


# ==========================================
# STEP 1: EXTRACT
# ==========================================

# --- Static Variable ---

class ExtractError(RuntimeError):
    pass

# --- Step Method ---

def extract_docx(docx_path: Path, work_dir: Path) -> Path:
    """Copy `docx_path` into `work_dir`, convert it, return the output .md path."""
    quarto = shutil.which("quarto")
    if not quarto:
        raise ExtractError("quarto not found on PATH")

    work_dir.mkdir(parents=True, exist_ok=True)
    local_docx = work_dir / docx_path.name
    shutil.copy2(docx_path, local_docx)

    out_md = work_dir / f"{docx_path.stem}.md"
    cmd = [
        quarto, "pandoc", local_docx.name,
        "-f", "docx", "-t", "markdown",
        "--extract-media=.",
        "-o", out_md.name,
    ]
    # cwd=work_dir so --extract-media=. drops media/ beside the output file.
    proc = subprocess.run(cmd, cwd=work_dir, capture_output=True, text=True)
    if proc.returncode != 0:
        raise ExtractError(
            f"quarto pandoc failed ({proc.returncode}): {proc.stderr.strip()[:500]}")
    if not out_md.exists():
        raise ExtractError(f"expected output missing: {out_md}")
    return out_md


# ==========================================
# STEP 2: SEGMENT
# ==========================================

# --- Static Variable ---

# Markdown structure regexes.
FENCE_RE = re.compile(r"^\s*(```|~~~)")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
_ATTR_RE = re.compile(r"\{[^}]*\}")
_SPAN_RE = re.compile(r"\[([^\]]*)\]\{[^}]*\}")
_SLUG_STRIP_RE = re.compile(r"[^a-z0-9]+")
_BOLD_LINE_RE = re.compile(r"^(?:\*\*(.+)\*\*|__(.+)__)$")

# Segmentation thresholds (chars for the size window; tokens/len for anchors/headings).
LLM_SEGMENT_MIN_CHARS = 1500
LLM_SEGMENT_MAX_CHARS = 60000
ANCHOR_MIN_TOKENS = 4
_MAX_HEADING_LEN = 80

@dataclass
class Chunk:
    index: int
    level: int          # 0 for the preamble, else heading depth 1..6
    heading: str        # cleaned heading text ("" for preamble)
    slug: str
    text: str           # full chunk text, including the heading line


# --- Utility Method ---

def _slugify(heading: str, index: int) -> str:
    text = _SPAN_RE.sub(r"\1", heading)      # [X]{.smallcaps} -> X
    text = _ATTR_RE.sub("", text)            # drop {#id .class}
    text = _SLUG_STRIP_RE.sub("-", text.lower()).strip("-")
    text = text or "section"
    return f"{index:02d}-{text}"


def _clean_heading(raw: str) -> str:
    text = _SPAN_RE.sub(r"\1", raw)
    text = _ATTR_RE.sub("", text)
    return text.strip()


def split_into_chunks(markdown: str) -> list[Chunk]:
    """Split `markdown` into heading-delimited chunks.

    Leading content before the first heading becomes a level-0 "preamble" chunk.
    """
    lines = markdown.splitlines()
    in_fence = False

    boundaries: list[tuple[int, int, str]] = []
    for i, line in enumerate(lines):
        if FENCE_RE.match(line):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        m = HEADING_RE.match(line)
        if m:
            boundaries.append((i, len(m.group(1)), m.group(2)))

    chunks: list[Chunk] = []
    idx = 0

    first = boundaries[0][0] if boundaries else len(lines)
    preamble = "\n".join(lines[:first]).strip("\n")
    if preamble.strip():
        chunks.append(Chunk(idx, 0, "", _slugify("preamble", idx), preamble))
        idx += 1

    for b, (start, level, raw_heading) in enumerate(boundaries):
        end = boundaries[b + 1][0] if b + 1 < len(boundaries) else len(lines)
        text = "\n".join(lines[start:end]).strip("\n")
        heading = _clean_heading(raw_heading)
        body = "\n".join(text.splitlines()[1:]).strip()
        if not heading and not body:
            continue
        chunks.append(Chunk(idx, level, heading, _slugify(raw_heading, idx), text))
        idx += 1

    return chunks


def count_headings(markdown: str) -> int:
    """Number of real `#`..`######` heading lines, ignoring fenced code blocks."""
    in_fence = False
    total = 0
    for line in markdown.splitlines():
        if FENCE_RE.match(line):
            in_fence = not in_fence
            continue
        if not in_fence and HEADING_RE.match(line):
            total += 1
    return total


def promote_pseudo_headings(markdown: str) -> tuple[str, int]:
    """Promote whole-line bold pseudo-headings to provisional `## ` headings.

    A line qualifies when its trimmed content is a single bold span, the inner text
    is non-empty and <= `_MAX_HEADING_LEN`, and contains no nested `**`/`__` (a
    two-span line is emphasis, not a heading). Fenced code blocks are skipped.
    Returns `(new_markdown, promoted_count)`.
    """
    lines = markdown.splitlines()
    in_fence = False
    promoted = 0
    for i, raw in enumerate(lines):
        if FENCE_RE.match(raw):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        match = _BOLD_LINE_RE.match(raw.strip())
        if not match:
            continue
        inner = (match.group(1) or match.group(2)).strip()
        if not inner or len(inner) > _MAX_HEADING_LEN or "**" in inner or "__" in inner:
            continue
        lines[i] = f"## {inner}"
        promoted += 1
    return "\n".join(lines), promoted


def _anchor_regex(tokens: list[str]) -> re.Pattern:
    """Whitespace-flexible, case-insensitive regex matching `tokens` in order."""
    return re.compile(r"\s+".join(re.escape(t) for t in tokens), re.IGNORECASE)


def _find_anchor(text: str, start: int, tokens: list[str]) -> int | None:
    """Offset of the best occurrence of `tokens` in `text[start:]`, or None.

    Tries the full token list, then successively shorter prefixes down to
    `ANCHOR_MIN_TOKENS` (so a trailing paraphrase still anchors). At each prefix
    length, prefers a match at a line start (offset 0 or preceded by a newline);
    if none is line-initial (e.g. a single giant paragraph), takes the first match.
    """
    for k in range(len(tokens), ANCHOR_MIN_TOKENS - 1, -1):
        matches = list(_anchor_regex(tokens[:k]).finditer(text, start))
        if not matches:
            continue
        for m in matches:
            if m.start() == 0 or text[m.start() - 1] == "\n":
                return m.start()
        return matches[0].start()
    return None


def _heading_core_pattern(heading: str) -> str | None:
    """Regex fragment matching an inline rendering of `heading`, or None if empty.

    Matches the heading tokens (whitespace-flexible) wrapped in an optional list
    marker ("A.", "1)", "(a)") and optional surrounding bold/italic markers, with
    trailing separator punctuation. Used to detect an inline title dangling on
    either side of a spliced heading.
    """
    tokens = heading.split()
    if not tokens:
        return None
    return (
        r"(?:[(\[]?[A-Za-z0-9]{1,3}[.)\]]\s*)?"   # optional list marker: A. 1) (a)
        r"[*_]{0,3}\s*"                            # optional opening emphasis
        + r"\s+".join(re.escape(t) for t in tokens)
        + r"[*_]{0,3}"                             # optional closing emphasis
        r"[\s:;.,·—–-]*"                           # trailing separators
    )


def _strip_trailing_heading(text: str, heading: str) -> str:
    """Remove an inline title dangling at the END of `text` (see _heading_core).

    When the anchor points at the body, the source's inline title lands at the end
    of the preceding slice ("...proyek Wikipedia: Sejarah"); trim it.
    """
    core = _heading_core_pattern(heading)
    if core is None:
        return text
    return re.sub(r"\s*" + core + r"$", "", text, flags=re.IGNORECASE)


def _leading_heading_len(text: str, heading: str) -> int:
    """Length of an inline title at the START of `text`, or 0.

    When the anchor INCLUDES the title, the body begins with it ("Perkembangan
    jumlah artikel: Sejak resmi...") and would duplicate the spliced heading; report
    how many chars to skip so the body starts after it.
    """
    core = _heading_core_pattern(heading)
    if core is None:
        return 0
    m = re.match(r"\s*" + core, text, flags=re.IGNORECASE)
    return m.end() if m else 0


def segment_with_llm(markdown: str, client, api_key: str, model: str) -> str:
    """Insert `## ` headings into a heading-less document at LLM-identified anchors.

    The model returns only headings + verbatim opening anchors; the body is never
    regenerated. Each anchor is located (with prefix fallback + line-start
    preference) in the text after the previous insertion, keeping sections ordered.
    On any failure or when nothing can be located, returns `markdown` unchanged.
    """
    def _validate(content: str) -> None:
        sections = json.loads(content).get("sections")
        if not isinstance(sections, list):
            raise ValueError("response has no 'sections' list")
        for entry in sections:
            if not (isinstance(entry, dict)
                    and isinstance(entry.get("heading"), str)
                    and isinstance(entry.get("anchor"), str)):
                raise ValueError("bad section entry shape")

    try:
        content = chat_completion(
            client,
            api_key,
            model,
            [
                {"role": "system", "content": SEGMENT_PROMPT},
                {"role": "user", "content": markdown},
            ],
            temperature=0.0,
            response_format={"type": "json_object"},
            validate=_validate,
        )
    except LLMError as e:
        _info(f"  ! segmentation failed, keeping one chunk: {e}")
        return markdown

    sections = json.loads(content)["sections"]
    if not sections:
        return markdown

    pieces: list[str] = []
    cursor = 0
    inserted = 0
    for entry in sections:
        heading = " ".join(entry["heading"].split()).lstrip("#").strip()
        if not heading:
            continue
        pos = _find_anchor(markdown, cursor, entry["anchor"].split())
        if pos is None:
            _info(f"  ! anchor not found, skipping heading: {heading!r}")
            continue
        # The inline source title lands either at the end of the preceding slice or
        # the start of the body; strip both so it isn't duplicated beside the heading.
        before = _strip_trailing_heading(markdown[cursor:pos], heading)
        pieces.append(before)
        pieces.append(f"\n\n## {heading}\n\n")
        cursor = pos + _leading_heading_len(markdown[pos:], heading)
        inserted += 1
    if inserted == 0:
        return markdown
    pieces.append(markdown[cursor:])
    return "".join(pieces)

# --- Step Method ---

def derive_chunks(text: str, client, api_key: str, model: str,
                  allow_llm: bool = True) -> list:
    """Make implicit structure explicit, then split into chunks.

    Gate: only engage the pre-pass when the document has fewer than 2 real headings.
    Step A (deterministic) promotes standalone bold lines. If `allow_llm` and the
    result is still a single chunk within the size window, Step B asks the LLM to
    segment it. With `allow_llm=False` (no-LLM run) only the deterministic Step A
    runs, so `client`/`api_key` may be None. Prints a one-line summary of where the
    boundaries came from. Returns the chunk list.
    """
    real_headings = count_headings(text)
    promoted = 0
    used_llm = False

    if real_headings < 2:
        text, promoted = promote_pseudo_headings(text)

    chunks = split_into_chunks(text)

    if allow_llm and len(chunks) <= 1 and LLM_SEGMENT_MIN_CHARS < len(text) <= LLM_SEGMENT_MAX_CHARS:
        segmented = segment_with_llm(text, client, api_key, model)
        used_llm = segmented != text
        if used_llm:
            text = segmented
            chunks = split_into_chunks(text)

    _info(
        f"Boundaries: {len(chunks)} chunk(s) "
        f"(real={real_headings}, promoted={promoted}, llm_segmented={used_llm})"
    )
    return chunks


# ==========================================
# STEP 3: CLEAN
# ==========================================

# --- Static Variable ---

# Cleaner: reject output grown far beyond the input as likely hallucinated.
_SUMMARY_DELIM = "---SUMMARY---"
_MAX_GROWTH_RATIO = 1.5
_MAX_GROWTH_SLACK = 200

class CleanerError(RuntimeError):
    """Raised when a chunk cannot be cleaned after all retries."""

@dataclass
class Section:
    index: int
    level: int
    heading: str
    slug: str
    file: str
    summary: str = ""  # from the cleaning call; kept in structure.json for debugging


# --- Step Method ---

def clean_chunk(
    chunk: Chunk,
    client: httpx.Client,
    api_key: str,
    model: str = DEFAULT_MODEL,
    max_retries: int = 3,
) -> tuple[str, str]:
    """Return (cleaned Markdown, 1-2 sentence summary) for `chunk`.

    The summary is "" when the model omits the delimiter. Raises `CleanerError`
    after retries are exhausted (including when the output grew so much it looks
    hallucinated -- that rejection is retried, then surfaces as a failure).
    """

    def _reject_if_hallucinated(content: str) -> None:
        cleaned = content.partition(_SUMMARY_DELIM)[0].strip()
        limit = len(chunk.text) * _MAX_GROWTH_RATIO + _MAX_GROWTH_SLACK
        if len(cleaned) > limit:
            raise LLMError(
                f"output grew {len(chunk.text)} -> {len(cleaned)} chars; "
                "likely hallucinated content"
            )

    try:
        content = chat_completion(
            client,
            api_key,
            model,
            [
                {"role": "system", "content": CLEAN_SYSTEM_PROMPT},
                {"role": "user", "content": chunk.text},
            ],
            temperature=0.2,
            max_retries=max_retries,
            validate=_reject_if_hallucinated,
        )
    except LLMError as e:
        raise CleanerError(str(e)) from e

    cleaned, delim, summary = content.partition(_SUMMARY_DELIM)
    return cleaned.strip(), (summary.strip() if delim else "")


def build_manifest(chunks: list[Chunk], summaries: dict[int, str] | None = None) -> list[Section]:
    summaries = summaries or {}
    return [
        Section(c.index, c.level, c.heading, c.slug, f"{c.slug}.md",
                summaries.get(c.index, ""))
        for c in chunks
    ]


def write_manifest(sections: list[Section], path: Path) -> None:
    data = {"sections": [s.__dict__ for s in sections]}
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


# ==========================================
# STEP 4: ASSEMBLE
# ==========================================

# --- Static Variable ---

HEADING_PREFIX_RE = re.compile(r"^(#{1,6})(\s+.*)$")  # keeps whitespace after the hashes
_IMG_RE = re.compile(r"(!\[[^\]]*\]\()([^)\s]+)(\))(\{[^}]*\})?")
# A fragment-only link: [text](#anchor). Leading '!' marks an image -> skip.
_ANCHOR_LINK_RE = re.compile(r"(!?)\[([^\]]*)\]\(#([^)\s]+)\)")
# Inline markdown to strip before deriving a heading id (pandoc drops formatting).
_INLINE_LINK_RE = re.compile(r"\[([^\]]*)\]\([^)]*\)")

@dataclass
class PageGroup:
    """One final page: an H1 section plus everything under it."""
    title: str
    parts: list[str]     # body texts, in order
    indexes: list[int]   # section indexes contributing to this page


# --- Utility Method ---

def _manifest_payload(sections: list[Section]) -> str:
    entries = [
        {"index": s.index, "level": s.level, "heading": s.heading, "file": s.file}
        for s in sections
    ]
    return json.dumps({"sections": entries}, ensure_ascii=False)


def _yaml_front_matter(fields: dict[str, str | None]) -> str:
    """Render a YAML front matter block. JSON string escaping is valid YAML."""
    lines = ["---"]
    for key, value in fields.items():
        if value:
            lines.append(f"{key}: {json.dumps(str(value), ensure_ascii=False)}")
    lines.append("---")
    return "\n".join(lines) + "\n\n"


def _apply_heading(text: str, level: int, heading: str) -> str:
    """Rewrite the first heading line of `text` to the given level/heading.

    Level 0 (preamble) has no heading line -> returned unchanged.
    """
    if level < 1:
        return text
    new_line = f"{'#' * level} {heading}".rstrip()
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if HEADING_RE.match(line):
            lines[i] = new_line
            return "\n".join(lines)
    return new_line + "\n\n" + text


def _is_remote(path: str) -> bool:
    return path.startswith(("http://", "https://", "data:", "//"))


def _parse_descriptions(content: str, expected: int) -> list[str]:
    """Parse the model's JSON reply into a list of `expected` description strings.

    Raises `ValueError` if the JSON is malformed or the shape/length is wrong, so
    a bad response is retried by chat_completion rather than silently accepted.
    """
    descriptions = json.loads(content).get("descriptions")
    if (
        isinstance(descriptions, list)
        and len(descriptions) == expected
        and all(isinstance(d, str) for d in descriptions)
    ):
        return descriptions
    raise ValueError(f"bad shape: expected {expected} strings")


def heading_id(text: str) -> str:
    """Pandoc-style auto identifier for a heading, matching Pandoc's cross-ref ids.

    Strips inline formatting, drops disallowed punctuation, lowercases, joins words
    with hyphens, and removes any leading non-letters (so "1. Overview" -> "overview").
    Falls back to "section" when nothing is left. Must mirror the ids Pandoc emitted
    in the source's link anchors so the two line up.
    """
    stripped = _INLINE_LINK_RE.sub(r"\1", text)          # [t](u) -> t
    stripped = stripped.replace("`", "").replace("*", "")  # code / emphasis markers
    kept = [
        ch.lower() if (ch.isalnum() or ch in "_-.") else " " if ch.isspace() else ""
        for ch in stripped
    ]
    # Drop pure-punctuation tokens (e.g. a "---" standing in for an em dash, which
    # Pandoc removes) so runs of separators don't leak into the id.
    tokens = [t for t in "".join(kept).split() if any(c.isalnum() for c in t)]
    ident = "-".join(tokens)
    i = 0
    while i < len(ident) and not ident[i].isalpha():
        i += 1
    return ident[i:] or "section"


def _group_heading(group: PageGroup) -> str:
    """The first `# ...` heading text in a group's body, else its title.

    Used to match `--skip-toc`: the index group's title is the doc title, but its
    body may open with the TOC heading (e.g. 'Daftar Isi'), so we match on the body.
    """
    for part in group.parts:
        for line in part.splitlines():
            m = HEADING_RE.match(line)
            if m:
                return m.group(2).strip()
    return group.title


# --- Step Method ---

def evaluate_structure(
    sections: list[Section],
    client: httpx.Client,
    api_key: str,
    model: str,
) -> dict:
    """Ask the LLM to correct levels/headings. Returns the parsed JSON response.

    The model returns only a delta ("changes"); we overlay it onto the full section
    list locally. Falls back to the original structure (no reorder) if the call or
    parse fails after retries.
    """
    full = [
        {"index": s.index, "level": s.level, "heading": s.heading, "file": s.file}
        for s in sections
    ]
    fallback = {
        "title": None,
        "sections": full,
        "suggested_order": None,
        "reorder_reason": None,
    }

    def _log_retry(attempt: int, err: Exception) -> None:
        _info(f"  ! structure evaluation attempt {attempt + 1} failed, retrying: {err}")

    try:
        content = chat_completion(
            client,
            api_key,
            model,
            [
                {"role": "system", "content": STRUCTURE_PROMPT},
                {"role": "user", "content": _manifest_payload(sections)},
            ],
            temperature=0.0,
            response_format={"type": "json_object"},
            validate=json.loads,
            on_retry=_log_retry,
        )
    except LLMError as e:
        _info(f"  ! structure evaluation failed, keeping original outline: {e}")
        return fallback

    result = json.loads(content)
    by_index = {e["index"]: e for e in full}
    for change in result.get("changes") or []:
        entry = by_index.get(change.get("index"))
        if entry is None:
            continue
        if isinstance(change.get("level"), int) and entry["level"] > 0:
            entry["level"] = change["level"]
        if change.get("heading"):
            entry["heading"] = str(change["heading"])
    result["sections"] = full
    return result


def normalize_levels(sections: list[dict]) -> list[dict]:
    """Deterministically fix heading-level jumps, in place.

    A heading may be at most one level deeper than the previous one (H1 -> H3
    becomes H1 -> H2). The first heading becomes H1. Level 0 (preamble) is kept.
    Runs after the LLM evaluation as a safety net, so the split is sane even when
    the model call failed or returned bad levels.
    """
    prev = 0
    for sec in sections:
        if sec["level"] < 1:
            continue
        if sec["level"] > prev + 1:
            sec["level"] = prev + 1
        prev = sec["level"]
    return sections


def group_by_h1(
    chunk_dir: Path,
    sections: list[dict],
    order: list[int] | None = None,
    doc_title: str | None = None,
) -> list[PageGroup]:
    """Group sections into pages: a new page starts at every titled level-1 heading.

    The first group is always the index page (`doc_title`), collecting the
    preamble and anything before the first H1 -- it may be empty. A level-1
    section with no heading text (e.g. a docx TOC field) starts no page of its
    own -- it has no title to give one -- and is folded into the current page
    instead, with its own blank heading line dropped.
    """
    by_index = {s["index"]: s for s in sections}
    sequence = order if order else [s["index"] for s in sections]

    groups = [PageGroup(doc_title or "Overview", [], [])]
    for idx in sequence:
        sec = by_index[idx]
        text = (chunk_dir / sec["file"]).read_text(encoding="utf-8").strip("\n")
        text = _apply_heading(text, sec["level"], sec["heading"])
        if sec["level"] == 1 and sec["heading"]:
            groups.append(PageGroup(sec["heading"], [text], [idx]))
        else:
            if sec["level"] == 1:
                text = "\n".join(text.splitlines()[1:]).lstrip("\n")
            groups[-1].parts.append(text)
            groups[-1].indexes.append(idx)
    return groups


def describe_pages(
    pages: list[dict],
    client: httpx.Client,
    api_key: str,
    model: str,
) -> list[str]:
    """`pages`: [{"title": str, "summaries": [str, ...]}, ...] in page order.

    Returns one description per page. On failure, falls back to each page's
    first section summary ("" when a page has none).
    """
    fallback = [p["summaries"][0] if p["summaries"] else "" for p in pages]
    try:
        content = chat_completion(
            client,
            api_key,
            model,
            [
                {"role": "system", "content": DESCRIBE_PROMPT},
                {"role": "user", "content": json.dumps(pages, ensure_ascii=False)},
            ],
            response_format={"type": "json_object"},
            validate=lambda c: _parse_descriptions(c, len(pages)),
        )
    except LLMError as e:
        _info(f"  ! page descriptions failed, using first summaries: {e}")
        return fallback
    return _parse_descriptions(content, len(pages))


def drop_toc_page(
    groups: list[PageGroup],
    descriptions: list[str] | None,
    heading: str,
) -> tuple[list[PageGroup], list[str] | None, bool]:
    """Remove the first group whose heading equals `heading` (case-insensitive).

    Keeps `descriptions` aligned with the surviving groups. Returns
    (groups, descriptions, dropped) where `dropped` is False if nothing matched.
    """
    key = heading.strip().casefold()
    kept: list[PageGroup] = []
    kept_desc: list[str] | None = None if descriptions is None else []
    dropped = False
    for i, group in enumerate(groups):
        if not dropped and _group_heading(group).strip().casefold() == key:
            dropped = True
            continue
        kept.append(group)
        if kept_desc is not None:
            kept_desc.append(descriptions[i] if i < len(descriptions) else "")
    return kept, kept_desc, dropped


def _strip_leading_h1(body: str) -> str:
    """Drop a leading `# ` heading (and the blank line after it) from `body`.

    Only a level-1 heading is removed; `## ` and deeper are left untouched.
    """
    lines = body.split("\n")
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and lines[i].startswith("# "):
        del lines[i]
        while i < len(lines) and not lines[i].strip():
            del lines[i]
    return "\n".join(lines)


def split_by_h1(
    target_dir: Path,
    groups: list[PageGroup],
    front_matter_extra: dict[str, str | None] | None = None,
    descriptions: list[str] | None = None,
    write_meta: bool = True,
    index_page: bool = True,
    emit_front_matter: bool = True,
) -> list[Path]:
    """Write `groups` as numbered `xx-slug.md` pages, plus (when `write_meta`)
    `meta.json`, following the docs-site standard.

    With `index_page` (default), the first group is written as `00-index.md` (the
    landing page, kept even when empty) and the rest as `01-slug.md`, ... . With
    `index_page=False` every group is numbered `00-slug.md`, `01-slug.md`, ... by
    its own title and no dedicated index is written -- used when `--skip-toc`
    removes the table-of-contents page.

    When `emit_front_matter` (the Fumadocs layout), each file starts with a YAML
    front matter block (title, optional description from `descriptions` -- same
    order/length as `groups` -- plus `front_matter_extra`); the redundant leading
    `# <title>` is dropped from the body since the title already lives in the front
    matter. Otherwise (the flat layout) no front matter is written and the body
    opens with `# <title>`. `meta.json` lists the pages in order for the sidebar
    (Fumadocs-specific; skipped for a flat layout). Returns the written paths in order.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    def _write(path: Path, group: PageGroup, description: str | None) -> None:
        body = "\n\n".join(p for p in group.parts if p.strip())
        if emit_front_matter:
            fm = {"title": group.title, "description": description,
                  **(front_matter_extra or {})}
            body = _strip_leading_h1(body)
            path.write_text(_yaml_front_matter(fm) + body + "\n", encoding="utf-8")
        else:
            stripped = body.lstrip()
            if not (stripped.startswith("# ") or stripped == "#" or stripped.startswith("#\n")):
                body = f"# {group.title}\n\n{body}" if body else f"# {group.title}"
            path.write_text(body + "\n", encoding="utf-8")
        written.append(path)

    descriptions = descriptions or []

    def _desc(i: int) -> str | None:
        return descriptions[i] if i < len(descriptions) else None

    if index_page:
        _write(target_dir / "00-index.md", groups[0], _desc(0))
        for n, group in enumerate(groups[1:], start=1):
            _write(target_dir / f"{n:02d}-{slugify(group.title)}.md", group, _desc(n))
    else:
        for n, group in enumerate(groups):
            _write(target_dir / f"{n:02d}-{slugify(group.title)}.md", group, _desc(n))

    if write_meta and groups:
        meta = {"title": groups[0].title,
                "pages": [p.stem for p in written]}
        (target_dir / "meta.json").write_text(
            json.dumps(meta, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return written


def fix_file_headings(path: Path, start_level: int = 0) -> bool:
    """Clamp heading-level jumps inside one final file (fence-aware).

    Every heading may be at most one level deeper than the previous one
    (H1 -> H3 becomes H1 -> H2). `start_level` seeds the previous depth: 0 for a
    flat page whose body opens with its own H1, or 1 for a Fumadocs page whose H1
    title lives in the front matter (so a leading H2 is allowed, not promoted).
    Per-file safety net on top of the global normalization, since grouping/reordering
    can reintroduce jumps. Returns True if the file was rewritten.
    """
    lines = path.read_text(encoding="utf-8").splitlines()
    in_fence = False
    prev = start_level
    changed = False
    for i, line in enumerate(lines):
        if FENCE_RE.match(line):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        m = HEADING_PREFIX_RE.match(line)
        if not m:
            continue
        level = len(m.group(1))
        if level > prev + 1:
            level = prev + 1
            lines[i] = "#" * level + m.group(2)
            changed = True
        prev = level
    if changed:
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return changed


def reclaim_images(
    md_path: Path,
    source_dir: Path,
    assets_dir: Path,
    link_prefix: str,
) -> tuple[int, list[str]]:
    """Copy images referenced by `md_path` into `assets_dir` and fix links.

    `source_dir` is the directory the original links are relative to (the input
    file's folder). Links are rewritten to `{link_prefix}/{name}` (e.g.
    /images/topic/img.png). Returns (copied_count, missing_link_list).
    """
    text = md_path.read_text(encoding="utf-8")
    copied: set[str] = set()
    missing: list[str] = []

    def repl(m: re.Match) -> str:
        prefix, path, suffix = m.group(1), m.group(2), m.group(3)
        if _is_remote(path):
            return f"{prefix}{path}{suffix}"
        src = (source_dir / path).resolve()
        name = src.name
        if not src.exists():
            missing.append(path)
            return f"{prefix}{path}{suffix}"
        assets_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, assets_dir / name)
        copied.add(name)
        return f"{prefix}{link_prefix.rstrip('/')}/{name}{suffix}"

    new_text = _IMG_RE.sub(repl, text)
    if new_text != text:
        md_path.write_text(new_text, encoding="utf-8")
    return len(copied), missing


def _page_anchor_index(pages: list[Path]) -> dict[str, str]:
    """Map every heading's id -> the page filename that contains it (flat layout).

    Ids are derived with `heading_id`, matching the anchors Pandoc put in the source
    links. Duplicate ids are suffixed `-1`, `-2`, ... like Pandoc, so links to a
    repeated heading still resolve. Fenced code blocks are skipped.
    """
    index: dict[str, str] = {}
    seen: dict[str, int] = {}
    for page in pages:
        in_fence = False
        for line in page.read_text(encoding="utf-8").splitlines():
            if FENCE_RE.match(line):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            m = HEADING_RE.match(line)
            if not m:
                continue
            base = heading_id(m.group(2))
            n = seen.get(base, 0)
            seen[base] = n + 1
            anchor = base if n == 0 else f"{base}-{n}"
            index[anchor] = page.name
    return index


def link_pages(pages: list[Path]) -> tuple[int, list[str]]:
    """Point cross-file fragment links at the page that owns the target heading.

    For each `[text](#anchor)`, if the heading lives on another page rewrite it to
    `[text](that-page.md#anchor)`; same-page links and image links are left alone.
    Anchors with no matching heading are left unchanged and collected for reporting.
    Returns (rewritten_count, unresolved_anchor_list).
    """
    index = _page_anchor_index(pages)
    rewritten = 0
    unresolved: list[str] = []

    for page in pages:
        def repl(m: re.Match) -> str:
            nonlocal rewritten
            bang, text, anchor = m.group(1), m.group(2), m.group(3)
            if bang:                       # image, not a link
                return m.group(0)
            target = index.get(anchor)
            if target is None:
                unresolved.append(anchor)
                return m.group(0)
            if target == page.name:        # same page: bare fragment already works
                return m.group(0)
            rewritten += 1
            return f"[{text}]({target}#{anchor})"

        text = page.read_text(encoding="utf-8")
        new_text = _ANCHOR_LINK_RE.sub(repl, text)
        if new_text != text:
            page.write_text(new_text, encoding="utf-8")
    return rewritten, unresolved


# ==========================================
# ENTRYPOINT
# ==========================================

# --- Step Method ---

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Structure-aware Markdown cleaner.")
    p.add_argument("input", nargs="?", default=str(ROOT_DIR / "experiment" / "output.md"))
    p.add_argument("--layout", choices=["flat", "fuma"], default="flat",
                   help="output layout: 'flat' (all .md files + a media/ subfolder, "
                        "default) or 'fuma' (Fumadocs: content/docs/<topic>/ + "
                        "public/images/<topic>/ + meta.json). 'fuma' requires the LLM.")
    p.add_argument("--model-name", dest="model", default=default_model(),
                   help="model id (default: $LLM_MODEL or the built-in default)")
    p.add_argument("--model-endpoint", default=None,
                   help="LLM API base URL (default: $LLM_BASE_URL or the built-in default)")
    p.add_argument("--model-key", default=None,
                   help="LLM API key (default: $LLM_API_KEY / $OPENROUTER_API_KEY)")
    p.add_argument("--no-merge", action="store_true", help="skip final merge step")
    p.add_argument("--skip-toc", metavar="HEADING", default=None,
                   help="drop the page whose first heading equals HEADING (the table-of-"
                        "contents page); remaining pages are renumbered with no index")
    p.add_argument("--keep-work", action="store_true",
                   help="keep the per-run render/<uuid> workspace (default: delete)")
    p.add_argument("--allow-reorder", action="store_true",
                   help="apply the LLM's suggested section reorder (if any)")
    p.add_argument("--use-llm", action="store_true",
                   help="run the LLM steps (segmentation, cleaning, structure eval, "
                        "descriptions). Without it (or without LLM_API_KEY) the run is "
                        "a deterministic passthrough of the raw quarto extraction.")
    return p.parse_args()


def main() -> int:
    load_dotenv(ROOT_DIR / ".env")
    args = parse_args()

    # CLI overrides win over env/.env; utility.base_url()/api_key() read these back.
    if args.model_endpoint:
        os.environ["LLM_BASE_URL"] = args.model_endpoint
    if args.model_key:
        os.environ["LLM_API_KEY"] = args.model_key

    fuma = args.layout == "fuma"
    key = api_key()
    # 'fuma' layout mandates the LLM; --use-llm requests it; either way needs a key.
    if fuma and not key:
        print("ERROR: --layout fuma requires the LLM, but LLM_API_KEY is not set "
              "(see .env.example).", file=sys.stderr)
        return 1
    if args.use_llm and not key:
        print("WARNING: --use-llm set but LLM_API_KEY missing; running without the "
              "LLM (raw quarto passthrough).", file=sys.stderr)
    use_llm = (args.use_llm or fuma) and bool(key)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: input not found: {input_path}", file=sys.stderr)
        return 1

    # Fixed scratch base under the repo (like render.py); each run gets a fresh
    # UUID subfolder that must not pre-exist.
    work_base = ROOT_DIR / "render"
    work_base.mkdir(exist_ok=True)
    work_dir = work_base / str(uuid.uuid4())
    work_dir.mkdir()
    # Final output lands in the caller's working directory (like render.py),
    # grouped in a subfolder named after the input.
    target_dir = Path.cwd() / input_path.stem
    topic = slugify(input_path.stem)
    if fuma:
        pages_dir = target_dir / "content" / "docs" / topic
        images_dir = target_dir / "public" / "images" / topic
        image_link_prefix = f"/images/{topic}"
    else:
        pages_dir = target_dir
        images_dir = target_dir / "media"
        image_link_prefix = "media"

    # Banner + progress bar mirror render.py's rich UI.
    header_text = Text()
    header_text.append(input_path.name, style="bold white")
    header_text.append(f"\nLayout: {args.layout}", style="dim white")
    header_text.append(
        f"\nLLM: {'on (' + args.model + ')' if use_llm else 'off -- raw passthrough'}",
        style="dim white" if use_llm else "italic dim",
    )
    _console.print(Panel(header_text, title="[bold cyan]Qlon reverse[/bold cyan]",
                         border_style="cyan", padding=(0, 2)))

    progress_columns = (
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
    )
    # segment + clean are always run; extract is docx-only; merge adds assemble + finalize.
    is_docx = input_path.suffix.lower() == ".docx"
    total_steps = 2 + (1 if is_docx else 0) + (0 if args.no_merge else 2)

    failures = 0
    written: list[Path] = []
    copied = 0
    try:
        with Progress(*progress_columns, console=_console) as progress:
            task = progress.add_task("Starting reverse pipeline...", total=total_steps)

            # Step: .docx input -> extract to Markdown in a scratch subfolder first.
            if is_docx:
                progress.update(task, description=f"Extracting {input_path.name} -> Markdown (quarto pandoc)...")
                try:
                    input_path = extract_docx(input_path, work_dir / "extract")
                except ExtractError as e:
                    _console.print(f"ERROR: {e}", style="red", markup=False, highlight=False)
                    return 1
                progress.advance(task)

            text = input_path.read_text(encoding="utf-8")
            _info(f"Model: {args.model}" if use_llm
                  else "LLM disabled -- raw quarto passthrough (pass --use-llm to enable)")

            sections = []
            summaries: dict[int, str] = {}
            manifest_path = work_dir / "structure.json"

            # Step: segment. Step: clean (per-chunk sub-bar).
            if use_llm:
                with httpx.Client() as client:
                    progress.update(task, description="Deriving document structure...")
                    chunks = derive_chunks(text, client, key, args.model)
                    progress.advance(task)
                    _info(f"Work scratch: {work_dir}")

                    progress.update(task, description="Cleaning sections with the LLM...")
                    clean_task = progress.add_task("cleaning", total=len(chunks))
                    for i, chunk in enumerate(chunks):
                        try:
                            # Heading-only chunk: nothing to clean; calling the LLM here
                            # invites hallucinated filler. Pass through untouched.
                            body = chunk.text.split("\n", 1)[1] if "\n" in chunk.text else ""
                            if chunk.level >= 1 and not body.strip():
                                cleaned, summary = chunk.text, ""
                            else:
                                cleaned, summary = clean_chunk(chunk, client, key, model=args.model)
                        except CleanerError as e:
                            _info(f"  ! chunk {chunk.index} ({chunk.slug}) failed: {e}")
                            cleaned, summary = chunk.text, ""  # fall back to raw so nothing is lost
                            failures += 1
                        if summary:
                            summaries[chunk.index] = summary
                        (work_dir / f"{chunk.slug}.md").write_text(cleaned + "\n", encoding="utf-8")
                        sections = build_manifest(chunks[: i + 1], summaries)
                        write_manifest(sections, manifest_path)
                        progress.advance(clean_task)
                    progress.remove_task(clean_task)
                    progress.advance(task)
            else:
                progress.update(task, description="Deriving document structure...")
                chunks = derive_chunks(text, None, None, args.model, allow_llm=False)
                progress.advance(task)
                _info(f"Work scratch: {work_dir}")

                progress.update(task, description="Writing sections (raw passthrough)...")
                clean_task = progress.add_task("writing", total=len(chunks))
                for chunk in chunks:
                    (work_dir / f"{chunk.slug}.md").write_text(chunk.text + "\n", encoding="utf-8")
                    progress.advance(clean_task)
                progress.remove_task(clean_task)
                sections = build_manifest(chunks, summaries)
                write_manifest(sections, manifest_path)
                progress.advance(task)

            if not args.no_merge:
                # Step: assemble (evaluate structure, group into pages, describe).
                progress.update(task, description="Evaluating structure and grouping pages...")
                if use_llm:
                    with httpx.Client() as client:
                        evaluated = evaluate_structure(sections, client, key, args.model)
                        (work_dir / "structure.eval.json").write_text(
                            json.dumps(evaluated, indent=2, ensure_ascii=False), encoding="utf-8")

                        order = None
                        suggested = evaluated.get("suggested_order")
                        if suggested:
                            reason = evaluated.get("reorder_reason") or "(no reason given)"
                            if args.allow_reorder:
                                order = suggested
                                _info(f"Applying suggested reorder: {reason}")
                            else:
                                _info(f"NOTE: LLM suggests reordering sections -> {suggested}\n"
                                      f"      reason: {reason}\n"
                                      f"      keeping source order; rerun with --allow-reorder to apply.")

                        normalize_levels(evaluated["sections"])

                        groups = group_by_h1(work_dir, evaluated["sections"], order,
                                             doc_title=evaluated.get("title") or input_path.stem)

                        pages_payload = [
                            {"title": g.title,
                             "summaries": [summaries[i] for i in g.indexes if i in summaries]}
                            for g in groups
                        ]
                        descriptions = describe_pages(pages_payload, client, key, args.model)
                        (work_dir / "descriptions.json").write_text(
                            json.dumps([{"page": g.title, "description": d}
                                        for g, d in zip(groups, descriptions)],
                                       indent=2, ensure_ascii=False), encoding="utf-8")
                else:
                    manifest_sections = [s.__dict__ for s in sections]
                    normalize_levels(manifest_sections)
                    groups = group_by_h1(work_dir, manifest_sections, None,
                                         doc_title=input_path.stem)
                    descriptions = None

                index_page = True
                if args.skip_toc:
                    groups, descriptions, dropped = drop_toc_page(
                        groups, descriptions, args.skip_toc)
                    if dropped:
                        index_page = False  # no forced landing page once the TOC is removed
                        _info(f"Skipped TOC page: {args.skip_toc!r} (pages renumbered, no index)")
                    else:
                        print(f"  ! --skip-toc {args.skip_toc!r} matched no page heading; "
                              "nothing skipped", file=sys.stderr)
                progress.advance(task)

                # Step: finalize (write pages, clamp headings, fix links, reclaim images).
                progress.update(task, description="Writing pages, fixing links, reclaiming images...")
                written = split_by_h1(
                    pages_dir, groups,
                    front_matter_extra={
                        "last_updated": date.today().isoformat(),
                        "status": "draft",
                    },
                    descriptions=descriptions,
                    write_meta=fuma,
                    index_page=index_page,
                    emit_front_matter=fuma,
                )

                fixed = sum(fix_file_headings(p, start_level=1 if fuma else 0) for p in written)
                if fixed:
                    _info(f"Heading levels clamped in {fixed} file(s).")

                # Flat layout: each page is a standalone file, so cross-page anchor links
                # must name the target file. 'fuma' resolves #anchor within one route, so
                # its links are left as-is.
                if not fuma:
                    rewritten, unresolved = link_pages(written)
                    _info(f"Cross-file links: {rewritten} rewritten")
                    if unresolved:
                        uniq = sorted(set(unresolved))
                        print(f"  ! {len(uniq)} anchor(s) not matched to any page, "
                              f"left as-is: {uniq[:10]}{' ...' if len(uniq) > 10 else ''}",
                              file=sys.stderr)

                missing: list[str] = []
                for path in written:
                    c, m = reclaim_images(path, input_path.parent, images_dir,
                                          image_link_prefix)
                    copied += c
                    missing.extend(m)
                _info(f"Images: {copied} copied to {images_dir}")
                if missing:
                    print(f"  ! {len(missing)} image(s) not found, links left as-is: {missing}",
                          file=sys.stderr)
                progress.advance(task)
    finally:
        if args.keep_work:
            _info(f"Kept work scratch: {work_dir}")
        else:
            shutil.rmtree(work_dir, ignore_errors=True)

    if not args.no_merge and written:
        _console.print(f"\n[bold green]Done![/bold green] {len(written)} file(s) written.\n")
        notes = Text()
        notes.append("Output: ", style="bold yellow")
        notes.append(f"{pages_dir}\n")
        notes.append("Images: ", style="bold yellow")
        notes.append(f"{copied} copied to {images_dir}\n")
        for path in written:
            notes.append(f"  - {path.name}\n", style="dim")
        _console.print(Panel(notes, title="[bold]Reverse complete[/bold]",
                             border_style="green", padding=(1, 2)))

    if failures:
        print(f"Done with {failures} failed chunk(s) (raw text kept).", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
