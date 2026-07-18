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
from tqdm import tqdm

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
        print(f"  ! segmentation failed, keeping one chunk: {e}")
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
            print(f"  ! anchor not found, skipping heading: {heading!r}")
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

    print(
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
        print(f"  ! structure evaluation attempt {attempt + 1} failed, retrying: {err}")

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
        print(f"  ! structure evaluation failed, keeping original outline: {e}")
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
    """Group sections into pages: a new page starts at every level-1 heading.

    The first group is always the index page (`doc_title`), collecting the
    preamble and anything before the first H1 -- it may be empty.
    """
    by_index = {s["index"]: s for s in sections}
    sequence = order if order else [s["index"] for s in sections]

    groups = [PageGroup(doc_title or "Overview", [], [])]
    for idx in sequence:
        sec = by_index[idx]
        text = (chunk_dir / sec["file"]).read_text(encoding="utf-8").strip("\n")
        text = _apply_heading(text, sec["level"], sec["heading"])
        if sec["level"] == 1:
            groups.append(PageGroup(sec["heading"], [text], [idx]))
        else:
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
        print(f"  ! page descriptions failed, using first summaries: {e}")
        return fallback
    return _parse_descriptions(content, len(pages))


def split_by_h1(
    target_dir: Path,
    groups: list[PageGroup],
    front_matter_extra: dict[str, str | None] | None = None,
    descriptions: list[str] | None = None,
    write_meta: bool = True,
) -> list[Path]:
    """Write `groups` as `00-index.md` + one numbered `xx-slug.md` per H1 page,
    plus (when `write_meta`) `meta.json`, following the docs-site standard.

    Each file starts with YAML front matter (title, optional description from
    `descriptions` -- same order/length as `groups` -- plus `front_matter_extra`)
    followed by a body that always opens with `# <title>`. An empty index page is
    still written: every folder needs a landing page. `meta.json` lists the pages
    in order for the sidebar (Fumadocs-specific; skipped for a flat layout).
    Returns the written paths in order.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    def _write(path: Path, group: PageGroup, description: str | None) -> None:
        fm = {"title": group.title, "description": description,
              **(front_matter_extra or {})}
        body = "\n\n".join(p for p in group.parts if p.strip())
        if not body.lstrip().startswith("# "):
            body = f"# {group.title}\n\n{body}" if body else f"# {group.title}"
        path.write_text(_yaml_front_matter(fm) + body + "\n", encoding="utf-8")
        written.append(path)

    descriptions = descriptions or []

    def _desc(i: int) -> str | None:
        return descriptions[i] if i < len(descriptions) else None

    _write(target_dir / "00-index.md", groups[0], _desc(0))
    for n, group in enumerate(groups[1:], start=1):
        _write(target_dir / f"{n:02d}-{slugify(group.title)}.md", group, _desc(n))

    if write_meta:
        meta = {"title": groups[0].title,
                "pages": [p.stem for p in written]}
        (target_dir / "meta.json").write_text(
            json.dumps(meta, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    return written


def fix_file_headings(path: Path) -> bool:
    """Clamp heading-level jumps inside one final file (fence-aware).

    The body opens with H1; every later heading may be at most one level deeper
    than the previous one (H1 -> H3 becomes H1 -> H2). Per-file safety net on top
    of the global normalization, since grouping/reordering can reintroduce jumps.
    Returns True if the file was rewritten.
    """
    lines = path.read_text(encoding="utf-8").splitlines()
    in_fence = False
    prev = 0
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

    failures = 0
    try:
        # Step 1: .docx input -> extract to Markdown in a scratch subfolder first.
        if input_path.suffix.lower() == ".docx":
            print(f"Extracting {input_path.name} -> Markdown (quarto pandoc)")
            try:
                input_path = extract_docx(input_path, work_dir / "extract")
            except ExtractError as e:
                print(f"ERROR: {e}", file=sys.stderr)
                return 1

        text = input_path.read_text(encoding="utf-8")
        print(f"Model: {args.model}" if use_llm
              else "LLM disabled -- raw quarto passthrough (pass --use-llm to enable)")

        sections = []
        summaries: dict[int, str] = {}
        manifest_path = work_dir / "structure.json"

        if use_llm:
            with httpx.Client() as client:
                chunks = derive_chunks(text, client, key, args.model)
                print(f"Work scratch: {work_dir}")
                for i, chunk in enumerate(tqdm(chunks, desc="cleaning", unit="chunk")):
                    try:
                        # Heading-only chunk: nothing to clean; calling the LLM here
                        # invites hallucinated filler. Pass through untouched.
                        body = chunk.text.split("\n", 1)[1] if "\n" in chunk.text else ""
                        if chunk.level >= 1 and not body.strip():
                            cleaned, summary = chunk.text, ""
                        else:
                            cleaned, summary = clean_chunk(chunk, client, key, model=args.model)
                    except CleanerError as e:
                        tqdm.write(f"  ! chunk {chunk.index} ({chunk.slug}) failed: {e}")
                        cleaned, summary = chunk.text, ""  # fall back to raw so nothing is lost
                        failures += 1
                    if summary:
                        summaries[chunk.index] = summary
                    (work_dir / f"{chunk.slug}.md").write_text(cleaned + "\n", encoding="utf-8")
                    sections = build_manifest(chunks[: i + 1], summaries)
                    write_manifest(sections, manifest_path)
        else:
            chunks = derive_chunks(text, None, None, args.model, allow_llm=False)
            print(f"Work scratch: {work_dir}")
            for chunk in chunks:
                (work_dir / f"{chunk.slug}.md").write_text(chunk.text + "\n", encoding="utf-8")
            sections = build_manifest(chunks, summaries)
            write_manifest(sections, manifest_path)

        if not args.no_merge:
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
                            print(f"Applying suggested reorder: {reason}")
                        else:
                            print(f"NOTE: LLM suggests reordering sections -> {suggested}\n"
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

            written = split_by_h1(
                pages_dir, groups,
                front_matter_extra={
                    "last_updated": date.today().isoformat(),
                    "status": "draft",
                },
                descriptions=descriptions,
                write_meta=fuma,
            )

            fixed = sum(fix_file_headings(p) for p in written)
            if fixed:
                print(f"Heading levels clamped in {fixed} file(s).")

            copied = 0
            missing: list[str] = []
            for path in written:
                c, m = reclaim_images(path, input_path.parent, images_dir,
                                      image_link_prefix)
                copied += c
                missing.extend(m)
            print(f"Images: {copied} copied to {images_dir}")
            if missing:
                print(f"  ! {len(missing)} image(s) not found, links left as-is: {missing}",
                      file=sys.stderr)
            print(f"Final clean files ({len(written)}) -> {pages_dir}")
            for path in written:
                print(f"  - {path.name}")
    finally:
        if args.keep_work:
            print(f"Kept work scratch: {work_dir}")
        else:
            shutil.rmtree(work_dir, ignore_errors=True)

    if failures:
        print(f"Done with {failures} failed chunk(s) (raw text kept).", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
