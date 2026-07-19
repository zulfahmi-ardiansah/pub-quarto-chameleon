# Qlon Web Interface

A browser front-end for Qlon, built as a 3-step wizard (Flowbite + Tailwind):

1. **Content** — paste Markdown (e.g. an AI response) for a single document, *or*
   upload one or more `.md`/`.qmd` files (each becomes a chapter, ordered by filename),
   *or* drop a single `.zip` project bundling your chapter Markdown plus its referenced
   images (e.g. an `images/` folder). The zip is unpacked server-side: chapter md is
   collected from any depth, a single wrapper folder is stripped, and relative image
   refs resolve automatically. Zip and loose-file uploads are mutually exclusive.
   The paste editor has a **Paste from clipboard** button and a **Preview** button that
   opens the rendered Markdown in a modal, plus a help link showing how to copy a
   response from ChatGPT, Claude, Gemini, or Perplexity (screenshots in `static/`).
   Uploads use a drag-and-drop zone; dropped files show as a numbered, removable list
   reflecting chapter order.
2. **Document info** — cover fields, page header, TOC title, and template selection.
3. **Render** — review and download the rendered `.docx` (zipped with an `Image/`
   folder when the document contains diagrams).

It wraps the existing CLI pipeline (`script/main.py`) — each request runs the same
render in an isolated, auto-deleted job workspace under `web/jobs/`.

The interface is a two-pane "workbench": an ink sidebar (brand, vertical stepper,
theme + language controls) beside a working surface. On the Content step the pasted
Markdown renders live as a **paper-sheet preview** beside the editor (collapses to a
modal on small screens). The front-end pulls Tailwind, `marked`, and `DOMPurify` from
CDNs, so the browser needs internet access. The preview uses `marked` (sanitized by
`DOMPurify`); it's an approximation — final styling comes from Quarto + the template.

Static assets are cache-busted with a `?v=` token: `app.py` derives it from the
newest file mtime in `static/` (`asset_v`), with a manual fallback in the template.
This makes edits to CSS/JS show on a plain refresh with the mounted-volume workflow.

**Theme & language**: a top-bar toggle switches light/dark mode (respects the OS
preference by default) and the interface language between English and Bahasa
Indonesia. Both choices persist in `localStorage`. UI strings live in
`static/i18n.js`; the primary brand color (`#005db5`) is set in the Tailwind config
inside `templates/index.html`.

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
pip install -r requirements.txt -r web/requirements-web.txt
playwright install chromium     # first time only

python web/app.py               # dev server on http://localhost:5000
```

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
- **Render timeout** is 600s. Large documents with many diagrams may need more — adjust
  `RENDER_TIMEOUT` in `app.py` and the gunicorn `--timeout`.
- On failure, Quarto's error output is surfaced in the page so you can fix the source.
