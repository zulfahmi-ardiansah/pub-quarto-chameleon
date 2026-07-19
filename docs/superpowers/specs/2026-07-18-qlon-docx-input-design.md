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

Split the monolithic `main.py` into three modules: a pure router plus one
module per input type.

### 1. Rename the bundle → `script/docx_ingest.py`
Rename `script/docx.py` → `script/docx_ingest.py`. Removes the `docx` import
shadow so `python-docx` resolves again. Then **clean it to match the existing
`main.py` style**, structurally only — logic stays byte-for-byte equivalent:
- Drop the generated-bundle markers (`# ===== from script/xxx.py =====`) and the
  "single-file bundle / re-run bundle.py" note in the docstring.
- Rewrite the module docstring in the concise style `main.py` uses.
- Import the prompt constants from `config/prompt.py` (see §6) instead of
  defining them inline.
- Keep all function bodies, names, comments, and control flow unchanged. This is
  a de-bundling + rename, not a behavioral rewrite.

### 2. Extract the render pipeline → `script/yml_ingest.py`
Move the entire current contents of `script/main.py` (the yml → DOCX render
pipeline: all helpers + its `main()`) **verbatim** into `script/yml_ingest.py` —
a straight move, no logic changes, so the current working render behavior is
preserved exactly. Paths are unchanged (`Path(__file__).parent.parent` is still
the repo root, since the file stays in `script/`).

### 3. `script/main.py` becomes a pure router
`main.py` no longer contains render logic. It inspects the positional input and
delegates:

```python
import sys

def main() -> int:
    positional = [a for a in sys.argv[1:] if not a.startswith("-")]
    if positional and positional[0].lower().endswith(".docx"):
        from docx_ingest import main as run   # .docx -> Markdown extraction
    else:
        from yml_ingest import main as run     # .yml  -> DOCX render (default)
    return run()

if __name__ == "__main__":
    raise SystemExit(main())
```

Lazy imports keep each pipeline's heavy dependencies (playwright/python-docx vs
httpx/tqdm) off the other's code path. Each callee runs its own argparse over
`sys.argv`, so `qlon file.docx --use-llm --layout fuma -o out/` and
`qlon x.yml --custom t.docx` pass through untouched.

**Guard rationale:** dispatch checks the first *positional* arg only, never flag
values. The render path's `--custom template.docx` takes a `.docx` as a flag
value; matching "any arg ends in .docx" would misfire on it. Positional-only
avoids that.

**Default routing:** anything that is not a `.docx` positional (a `.yml`, or no
positional at all, e.g. `--test`) goes to the render path, preserving current
behavior exactly.

### 4. Defaults
All inherited from the bundle as-is — no override layer:
- LLM off (deterministic raw quarto passthrough) unless `--use-llm`.
- Flat layout (all `.md` + `media/` in one folder).
- Output to `<input folder>/<input name>/`.

### 5. bin wrappers
`bin/qlon.sh` and `bin/qlon.bat` unchanged — they already forward all args to
`main.py`, which is now the router. Dispatch lives in one place.

### 6. Prompts → `config/prompt.py`
The four large prompt constants (`CLEAN_SYSTEM_PROMPT`, `STRUCTURE_PROMPT`,
`DESCRIBE_PROMPT`, `SEGMENT_PROMPT`) move out of the bundle into
`config/prompt.py`, with their existing explanatory comments kept.
`docx_ingest.py` imports them. Since `config/` is not on `sys.path`, the import
is enabled by inserting the config dir onto `sys.path` at module load
(consistent with how `main.py` already treats `ROOT_DIR / "config"` as
`CONFIG_DIR`):

```python
ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR / "config"))
from prompt import (
    CLEAN_SYSTEM_PROMPT, STRUCTURE_PROMPT, DESCRIBE_PROMPT, SEGMENT_PROMPT,
)
```

## Resulting module layout

- `script/main.py` — pure router (input type → pipeline).
- `script/yml_ingest.py` — yml → DOCX render pipeline (moved from main.py).
- `script/docx_ingest.py` — docx → Markdown extraction (de-bundled + renamed).
- `config/prompt.py` — the LLM prompt constants used by `docx_ingest.py`.

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
