# Web Zip Upload — Design

**Date:** 2026-07-20
**Status:** Approved (pre-implementation)

## Goal

Let the Qlon web interface accept a `.zip` archive that bundles chapter Markdown and
its referenced images (e.g. an `images/` folder), so users can render a whole project
folder in one upload instead of selecting loose `.md`/`.qmd` files. No CLI change.

## Context

The render path today ([web/app.py](../../../web/app.py) `render()`):

1. Accepts pasted Markdown *or* one-plus loose `.md`/`.qmd` uploads (`files`).
2. Stages each upload into `job_dir/content/`, writes `config.yml`, shells out to
   `script/main.py`, returns the DOCX (zipped with `Image/` when diagrams exist).

The CLI's `copy_content` ([script/render.py](../../../script/render.py)) globs
`*.qmd`/`*.md` **only at the content-folder top level** (non-recursive). For each
chapter it calls `_copy_images(text, f.parent)`, which resolves every relative image
ref against the md's own directory and copies it into the render workspace at the
**ref path** (`content_dir / raw`). HTTP/`/`-absolute refs are skipped. Mermaid blocks
are rendered separately and are unaffected by staging.

Implication for zip intake: staged chapter md must land at `content/` top level, and
each referenced image must land at `content/<ref>` so the existing resolver finds it.

## Approach

Unified drop zone accepts `.zip`. When a single `.zip` is uploaded, the server
extracts and **auto-normalizes** it, then hands the normalized `content/` folder to the
unchanged downstream pipeline.

### Frontend — [templates/index.html](../../../web/templates/index.html), [web/static/](../../../web/static/)

- Drop zone / file input `accept` gains `.zip`; drag-over validation accepts it.
- A dropped `.zip` renders as a single "project archive" item, visually distinct from
  the numbered chapter list used for loose md.
- Zip and loose-md uploads are **mutually exclusive**: selecting a zip clears any loose
  list; adding loose md clears a selected zip. Paste mode unchanged.
- New i18n strings (English + Bahasa Indonesia) in `static/i18n.js` for the zip item
  label, the mutual-exclusion hint, and zip-specific error text.

### Backend plumbing — [web/app.py](../../../web/app.py) `render()`

- Among uploaded `files`, detect the single-`.zip` case. If present, call
  `stage_zip(archive, content_dir)` instead of the per-file `.md/.qmd` staging loop.
- A zip mixed with loose md, or more than one zip → 400 (clear message).
- Everything after staging (config.yml assembly, subprocess call, DOCX/zip response,
  `finally` cleanup) is unchanged.

### New module — `web/zip_intake.py`

`stage_zip(archive_path: Path, content_dir: Path) -> None` (raises `ZipIntakeError`
with a user-facing message on any rejection). Kept in its own module so `app.py` stays
focused and the logic is unit-testable without Flask.

Algorithm:

1. **Open + validate members.** Reject any entry whose normalized path is absolute,
   contains `..`, or carries a drive letter (zip-slip). Enforce a per-member and a
   total **uncompressed** size cap (zip-bomb guard) inside the existing 64 MB request
   cap. Reject encrypted entries.
2. **Extract** to a temp staging dir (inside `job_dir`, cleaned by the existing
   `finally`).
3. **Normalize wrapper.** If every entry shares one common top-level directory, treat
   that directory as the root (strip the wrapper). Otherwise root = staging dir.
4. **Find chapters.** All `*.md`/`*.qmd` at any depth under root, sorted by their
   path-relative name for deterministic chapter order.
5. **Stage chapters.** Copy each md to `content_dir/<staged-name>` at the top level.
   On basename collision across folders, disambiguate the staged name by prefixing a
   sanitized form of its relative folder, so ordering stays deterministic and stable.
6. **Stage images.** For each chapter, scan its relative image refs (same rule as the
   CLI: skip `http(s)://` and `/`-absolute). For each ref, resolve the source against
   the md's original directory; if it exists, copy it to `content_dir/<sanitized-ref>`
   (leading `../` and `/` stripped; result kept within `content_dir`). md text is left
   unchanged when its refs already resolve after staging. Unreferenced files are
   ignored — matching current pipeline behavior. Mermaid untouched.
7. **Empty result.** If no chapter md was found, raise `ZipIntakeError`.

### Error handling

`stage_zip` raises `ZipIntakeError(message)`; `render()` maps it to
`jsonify(error=message), 400`, same response shape as existing validation errors
(bad file type, no content). Covers: not a valid zip, unsafe entry, oversize/zip-bomb,
encrypted entry, no md found.

## Edge cases

- **Single wrapper folder** (`project/ch1.md`, `project/images/x.png`) → wrapper
  stripped, resolves normally.
- **md at root + `images/` subfolder** → refs like `images/x.png` copy to
  `content/images/x.png`.
- **Ref climbs out** (`../assets/x.png`) → source resolved against md's real parent;
  destination sanitized to stay within `content/`.
- **Basename collision** (`a/notes.md`, `b/notes.md`) → folder-prefixed staged names.
- **Unreferenced images** → dropped (pipeline only copies referenced images).
- **Zip-slip / absolute / drive-letter entries** → rejected before extraction.

## Testing

Unit tests for `stage_zip` (pytest, build zips in a tmp dir):

- Flat zip (md + `images/` at root) stages md at top level, image at `content/images/`.
- Single-wrapper zip strips the wrapper.
- Nested chapters at multiple depths are all found, ordered deterministically.
- Zip-slip entry (`../evil`, absolute path) is rejected.
- Zip with no md is rejected.
- Basename collision across folders yields distinct staged names, both present.
- Oversize / zip-bomb (total uncompressed over cap) is rejected.

## Out of scope

- CLI changes to `script/`.
- Rewriting md image refs beyond what staging requires.
- Nested archives, non-image assets, per-chapter config in the zip.

## Files

- `web/zip_intake.py` — new: `stage_zip`, `ZipIntakeError`.
- `web/app.py` — detect zip, call `stage_zip`, map error.
- `web/templates/index.html`, `web/static/*` — drop-zone `.zip`, zip item UI,
  mutual exclusion, i18n.
- `web/tests/` (or existing test location) — `stage_zip` unit tests.
- `web/README.md` — Content-step note on zip upload.
