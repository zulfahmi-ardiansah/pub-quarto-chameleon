"""Stage an uploaded project zip (chapter Markdown + referenced images) into a
render workspace's content folder.

The CLI render pipeline (``script/render.py``) globs ``*.md``/``*.qmd`` only at the
content-folder top level and resolves each chapter's relative image refs against that
same folder. This module normalizes an arbitrary zip layout to satisfy both: chapter
md is placed at the content-dir top level, and every referenced image is copied to the
path the (possibly rewritten) ref resolves to.

``stage_zip`` raises :class:`ZipIntakeError` — with a user-facing message — on any
rejected archive (not a zip, unsafe member, oversize/zip-bomb, encrypted, no md).
"""

import re
import shutil
import zipfile
from pathlib import Path, PurePosixPath

# Mirror the CLI's image-reference regex so we scan refs identically.
_IMG_REF = re.compile(r"!\[.*?\]\(([^)]+)\)")

MAX_MEMBER_BYTES = 50 * 1024 * 1024   # per-file uncompressed cap
MAX_TOTAL_BYTES = 200 * 1024 * 1024   # whole-archive uncompressed cap (zip-bomb guard)

_MARKDOWN_SUFFIXES = {".md", ".qmd"}


class ZipIntakeError(Exception):
    """Raised when an uploaded zip cannot be safely staged. Message is user-facing."""


def _safe_member_path(name: str) -> PurePosixPath | None:
    """Return the member's normalized relative path, or None for a directory entry.

    Raises ZipIntakeError for unsafe entries (absolute, drive-letter, or ``..`` escape).
    """
    if name.endswith("/"):
        return None  # directory entry — created implicitly when files are written
    posix = name.replace("\\", "/")
    pure = PurePosixPath(posix)
    if pure.is_absolute() or (len(posix) >= 2 and posix[1] == ":"):
        raise ZipIntakeError(f"Unsafe path in zip: {name!r}")
    if any(part == ".." for part in pure.parts):
        raise ZipIntakeError(f"Unsafe path in zip: {name!r}")
    return pure


def _extract(archive_path: Path, dest: Path) -> None:
    """Validate every member, then extract the archive to *dest* safely."""
    try:
        zf = zipfile.ZipFile(archive_path)
    except zipfile.BadZipFile as exc:
        raise ZipIntakeError("Uploaded file is not a valid zip archive.") from exc

    with zf:
        total = 0
        for info in zf.infolist():
            if info.flag_bits & 0x1:
                raise ZipIntakeError("Encrypted zip archives are not supported.")
            rel = _safe_member_path(info.filename)
            if rel is None:
                continue
            if info.file_size > MAX_MEMBER_BYTES:
                raise ZipIntakeError(f"File too large in zip: {info.filename!r}.")
            total += info.file_size
            if total > MAX_TOTAL_BYTES:
                raise ZipIntakeError("Zip contents exceed the uncompressed size limit.")

        for info in zf.infolist():
            rel = _safe_member_path(info.filename)
            if rel is None:
                continue
            target = dest / Path(*rel.parts)
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)


def _strip_wrapper(root: Path) -> Path:
    """If every file lives under one common top-level folder, return that folder."""
    top_levels = set()
    for f in root.rglob("*"):
        if f.is_file():
            top_levels.add(f.relative_to(root).parts[0])
    if len(top_levels) == 1:
        candidate = root / next(iter(top_levels))
        if candidate.is_dir():
            return candidate
    return root


def _sanitize_ref(raw: str) -> str:
    """Reduce a relative ref to a path that stays inside the content dir.

    Strips a leading ``/`` and any leading/interior ``..`` and ``.`` components.
    """
    parts = [p for p in PurePosixPath(raw.replace("\\", "/")).parts if p not in ("..", ".", "/")]
    return str(PurePosixPath(*parts)) if parts else raw


def _unique_name(rel: PurePosixPath, used: set[str]) -> str:
    """Pick a top-level staged filename for a chapter, disambiguating collisions.

    First choice is the basename; on collision, fold the parent folders into the name
    so both chapters survive with deterministic, distinct names.
    """
    name = rel.name
    if name not in used:
        used.add(name)
        return name
    folded = "-".join(rel.parts)  # e.g. a/notes.md -> a-notes.md
    candidate = folded
    n = 1
    while candidate in used:
        candidate = f"{PurePosixPath(folded).stem}-{n}{PurePosixPath(folded).suffix}"
        n += 1
    used.add(candidate)
    return candidate


def _rewrite_refs(text: str, mapping: dict[str, str]) -> str:
    """Replace image-ref paths per *mapping* (original -> new), preserving alt/title."""
    if not mapping:
        return text

    def repl(m: re.Match) -> str:
        inner = m.group(1)
        head, _, _tail = inner.partition(" ")
        raw = head
        new = mapping.get(raw)
        if new is None:
            return m.group(0)
        rest = inner[len(raw):]  # keep any title/attrs suffix verbatim
        prefix = m.group(0)[: m.group(0).rindex("(") + 1]
        return f"{prefix}{new}{rest})"

    return _IMG_REF.sub(repl, text)


def stage_zip(archive_path: Path, content_dir: Path) -> None:
    """Extract *archive_path* and stage its Markdown + images into *content_dir*.

    Raises :class:`ZipIntakeError` (user-facing message) on any rejected archive.
    """
    staging = content_dir.parent / "_zip"
    if staging.exists():
        shutil.rmtree(staging, ignore_errors=True)
    staging.mkdir(parents=True)

    _extract(archive_path, staging)
    root = _strip_wrapper(staging)
    root_resolved = root.resolve()

    chapters = sorted(
        (f for f in root.rglob("*") if f.is_file() and f.suffix.lower() in _MARKDOWN_SUFFIXES),
        key=lambda f: f.relative_to(root).as_posix(),
    )
    if not chapters:
        raise ZipIntakeError("No Markdown (.md/.qmd) files found in the zip.")

    used: set[str] = set()
    for md in chapters:
        rel = PurePosixPath(md.relative_to(root).as_posix())
        text = md.read_text(encoding="utf-8")
        mapping: dict[str, str] = {}
        seen: set[str] = set()
        for match in _IMG_REF.finditer(text):
            raw = match.group(1).split(maxsplit=1)[0]
            if raw in seen or raw.startswith(("http://", "https://", "/")):
                continue
            seen.add(raw)
            src = (md.parent / raw).resolve()
            # Containment guard: a ref like `../../../etc/passwd` resolves outside the
            # extracted project. Refuse to read (and exfiltrate via the output) anything
            # that escapes the staging root. This — not the scheme check above — is the
            # control that stops arbitrary file reads; symlinks resolve() away too.
            try:
                src.relative_to(root_resolved)
            except ValueError:
                continue
            if not src.is_file():
                continue
            clean = _sanitize_ref(raw)
            if not clean or clean.startswith(".."):
                continue
            dest = content_dir / clean
            dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(src, dest)
            if clean != raw:
                mapping[raw] = clean
        staged_name = _unique_name(rel, used)
        (content_dir / staged_name).write_text(_rewrite_refs(text, mapping), encoding="utf-8")
