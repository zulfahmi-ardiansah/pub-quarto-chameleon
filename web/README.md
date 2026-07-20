# Qlon Web Interface

A browser front-end for Qlon (Flowbite + Tailwind). It runs in **both directions**,
switched by the **DOCX → MD** toggle in the top bar:

- **Render** (Markdown → Word) — a 3-step wizard, the default.
- **Reverse** (Word → Markdown) — a 4-step wizard.

Both wrap the same CLI pipeline (`script/main.py`); each request runs in an isolated,
auto-deleted job workspace under `web/jobs/`.

## Render wizard (Markdown → Word)

1. **Content** — paste Markdown (e.g. an AI response) for a single document, *or*
   upload one or more `.md`/`.qmd` files (each becomes a chapter, ordered by filename),
   *or* drop a single `.zip` project bundling your chapter Markdown plus its referenced
   images (e.g. an `images/` folder). The zip is unpacked server-side: chapter md is
   collected from any depth, a single wrapper folder is stripped, and relative image
   refs resolve automatically. Zip and loose-file uploads are mutually exclusive.
   The paste editor has a **Paste from clipboard** button and a **Preview** button that
   opens the rendered Markdown in a modal, plus a help link showing how to copy a
   response from ChatGPT, Claude, Gemini, or Perplexity (screenshots in `static/img/`).
   Uploads use a drag-and-drop zone; dropped files show as a numbered, removable list
   reflecting chapter order.
2. **Details** — cover fields, page header, TOC title, and template selection.
3. **Review** — review and download the rendered `.docx` (zipped with an `Image/`
   folder when the document contains diagrams).

## Reverse wizard (Word → Markdown)

1. **Upload** — drop a single `.docx` (exactly one; multiple are rejected).
2. **Options** — output **layout** (`flat` = `.md` files + `media/`, or `fuma` =
   Fumadocs tree, which requires the LLM), a **Skip TOC heading** field, and an **AI
   cleanup (LLM)** toggle. Enabling the LLM (or picking `fuma`) reveals model name /
   endpoint / API key fields and an **apply suggested reorder** checkbox. The key is
   used once for the conversion and never stored.
3. **Review** — review and convert.
4. **Done** — download the produced pages + media as a `.zip`.

## Server layers

- `app.py` — Flask app + routes (thin controllers): parse the request, call the
  service, translate the result or a `RenderError` into an HTTP response. Routes:
  `GET /`, `GET /template/download/<name>`, `POST /render`, `POST /reverse`.
- `render.py` — `render_document`: orchestrates one render job end to end (workspace,
  staging, `config.yml`, CLI subprocess, packaging). Also defines `RenderError`,
  `RenderResult`, and `run_pipeline` (the shared CLI-subprocess runner).
- `reverse.py` — `reverse_document`: orchestrates one reverse job (save upload,
  `build_reverse_args`, CLI subprocess, zip the page/media folder). Reuses
  `run_pipeline` from `render.py`.
- `staging.py` — `stage_content`: routes pasted text, loose `.md`/`.qmd`, or a project
  `.zip` (safe extraction + normalization via `stage_zip`) into the content folder.
- `util.py` — pure helpers (`build_config`, `list_presets`, `package_output`,
  `package_folder`).
- `paths.py` — shared filesystem locations.

The pipeline is invoked as a subprocess (not imported) because `script/main.py` relies
on module-level globals, the working directory, and rich console output — shelling out
reuses the tested code path. A `.docx` positional routes it to reverse; a `config.yml`
positional to render.

## Front-end

The interface is a two-pane "workbench": an ink header (brand, per-wizard stepper,
mode / theme / language controls) above a working surface. The front-end pulls Tailwind,
`marked`, and `DOMPurify` from CDNs (and fonts from Google Fonts), so the browser needs
internet access. The paste-mode preview uses `marked` (sanitized by `DOMPurify`) and
opens in a modal; it's an approximation — final styling comes from Quarto + the template.

Scripts (`static/js/`): `i18n.js` (UI strings), `app.js` (page chrome: theme, language,
steppers, modals), `render.js` (render wizard), `reverse.js` (reverse wizard), `mode.js`
(the forward/reverse toggle, which flips `body.mode-reverse`).

Static assets are cache-busted with a `?v=` token: `app.py` derives it from the newest
file mtime in `static/` (`asset_v`), with a manual fallback in the template. This makes
edits to CSS/JS show on a plain refresh with the mounted-volume workflow.

**Theme & language**: a top-bar toggle switches light/dark mode (respects the OS
preference by default) and the interface language between English and Bahasa Indonesia.
Both choices persist in `localStorage`. UI strings live in `static/js/i18n.js`; the
primary brand color (`#005db5`) is set in the Tailwind config inside
`templates/index.html`.

## Run with Docker (recommended)

The image bundles Python, Quarto, and a headless Chromium (for Mermaid), so nothing
else is needed on the host.

```bash
docker compose up --build
```

Then open <http://localhost:5000>. (First build is large — Quarto + Chromium.)

`web/static/` and `web/templates/` are mounted as volumes, so edits to the CSS, JS,
or HTML show on a browser refresh — no rebuild or `docker compose down` needed.
(Changes to `app.py` still require a container restart.)

## Run locally

Requires Python 3.13, Quarto, and Playwright Chromium already installed and on PATH
(see the project README). From the repo root:

```bash
pip install -r web/requirements-web.txt
playwright install chromium     # first time only

python web/app.py               # dev server on http://localhost:5000
```

> `web/requirements-web.txt` is self-contained (it includes the pipeline deps) and
> installs on Linux; the repo-root `requirements.txt` is a frozen Windows env dump.

For a production-style server instead of the Flask dev server:

```bash
cd web
gunicorn -w 2 --timeout 600 -b 0.0.0.0:5000 app:app
```

## Notes

- **Chapter order** follows the uploaded filenames, sorted alphabetically. Prefix them
  (`01-intro.md`, `02-setup.md`, …) to control ordering.
- **Templates**: choose a built-in preset, or upload a custom `.docx` reference
  template (which takes precedence over the preset).
- **Timeouts**: render is 600s (`RENDER_TIMEOUT` in `render.py`), reverse is 900s
  (`REVERSE_TIMEOUT` in `reverse.py`) since the LLM path is slower. Adjust these and the
  gunicorn `--timeout` for large documents.
- **Reverse LLM**: the key is passed through to the CLI for a single conversion and
  never persisted. Without a key (LLM off), reverse is a deterministic Quarto passthrough.
- On failure, the CLI's error output is surfaced in the page so you can fix the source.
- Tests live in `web/tests/` (e.g. `test_staging.py`).
