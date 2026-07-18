# qlon `.docx` input → Markdown extraction

## Goal

When `qlon`'s input is a `.docx` file, run the Markdown-extraction pipeline
(the bundle currently at `script/docx.py`) instead of the yml → DOCX renderer.
Output is cleaned Markdown chapters + media, exactly as the bundle already
produces. No behavior change to the extraction itself.

## Problem being fixed

The bundle was dropped in as `script/docx.py`. Because the script directory is
first on `sys.path`, `docx.py` **shadows the `python-docx` package**, so
`main.py`'s `from docx import Document` resolves to the bundle and raises
`ImportError`. The normal `qlon <config.yml>` render path is broken as long as
`script/docx.py` exists. Confirmed:

```
ImportError: cannot import name 'Document' from 'docx'
  (D:\Repository\personal-quarto-chameleon\script\docx.py)
```

## Design

### 1. Rename the bundle
`script/docx.py` → `script/docx_ingest.py`. Content unchanged. Removes the
`docx` import shadow so `python-docx` resolves again.

### 2. Early dispatch in `main.py`
At the top of `main()`, before argparse and before the render pipeline runs,
detect a `.docx` **positional** input and delegate to the ingest pipeline:

```python
# A .docx positional input routes to the Markdown-extraction pipeline
# (script/docx_ingest.py) instead of the yml -> DOCX renderer.
positional = [a for a in sys.argv[1:] if not a.startswith("-")]
if positional and positional[0].lower().endswith(".docx"):
    from docx_ingest import main as ingest_main
    return ingest_main()
```

`ingest_main()` runs its own argparse over `sys.argv`, so
`qlon file.docx --use-llm --layout fuma -o out/` pass through untouched. Each
pipeline keeps its own flags and dependencies isolated.

**Guard rationale:** dispatch checks the first *positional* arg only, never flag
values. The render path's `--custom template.docx` takes a `.docx` as a flag
value; matching "any arg ends in .docx" would misfire on it. Positional-only
avoids that.

### 3. Defaults
All inherited from the bundle as-is — no override layer:
- LLM off (deterministic raw quarto passthrough) unless `--use-llm`.
- Flat layout (all `.md` + `media/` in one folder).
- Output to `<input folder>/<input name>/`.

### 4. bin wrappers
`bin/qlon.sh` and `bin/qlon.bat` unchanged — they already forward all args to
`main.py`; dispatch lives in one place.

## Flow

- `qlon x.yml [render flags]` → render path (unchanged).
- `qlon x.docx [ingest flags]` → extraction path.
- The positional input's extension is the only selector.

## Out of scope

- No round-trip (extracted Markdown is the final output, not re-rendered).
- No changes to the extraction bundle's own logic.
- No new subcommand or CLI surface.

## Verification

- `qlon <config.yml>` (or `--test`) still renders a DOCX — regression check that
  the rename unshadowed `python-docx`.
- `qlon sample.docx` produces `sample/` with Markdown + `media/`.
- `qlon --custom template.docx <config.yml>` still routes to the render path
  (positional guard holds).
