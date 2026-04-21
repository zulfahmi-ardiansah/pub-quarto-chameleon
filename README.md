# Qlon — Quarto Cameloon

![Version](https://img.shields.io/badge/version-alpha-orange) ![Python](https://img.shields.io/badge/python-3.13+-3776AB?logo=python&logoColor=white) ![Quarto](https://img.shields.io/badge/quarto-1.x-75AADB?logo=quarto&logoColor=white)

**Qlon** (short for *Quarto Cameloon*) is a command-line tool that converts Quarto Markdown (`.qmd`) into styled Word documents and PDFs.

Writing documentation in Word is tedious, hard to track, and doesn't play well with version control. Markdown solves all of that — and with AI now able to generate `.md` files directly from source code and logic, authoring docs has never been faster. The only friction left is delivery: most people still expect a `.docx` or `.pdf`. Qlon removes that friction.

You write in `.qmd`, describe your document in a YAML config file, and Qlon handles the rest — assembling the workspace, injecting metadata, rendering to Word via [Quarto](https://quarto.org/), replacing header placeholders, generating a PDF, and delivering the finished files to your working directory. No touching Quarto configs or Word templates required. Just like a Cameloon.

---

## How It Works

Each run gets its own isolated workspace at `render/<uuid>/`. Qlon stages all assets there, runs Quarto, post-processes the output, delivers the finished files to your working directory, then deletes that folder. The `render/` parent directory persists but stays empty between runs. Your source files are never modified.

When you run `qlon <config.yml>`, the following steps execute in order:

| # | Step | What happens |
|---|------|-------------|
| 1 | Set up workspace | Creates a UUID-named subfolder inside `render/` and copies the base Quarto config, cover page, and Word reference template into it. The template used is `basic.docx` by default, or the one selected via `--preset` / `--custom`. |
| 2 | Patch TOC title | *(Optional)* If `content.table.title` is set in your config, rewrites the title in the cover page frontmatter. |
| 3 | Copy content | Copies all `.qmd` chapter files from your content folder into the workspace, sorted alphabetically. |
| 4 | Configure Quarto | Injects cover metadata (title, subtitle, author, date) and the chapter list into the workspace `_quarto.yml`. Any `quarto:` overrides in your config are deep-merged on top. |
| 5 | Render | Runs `quarto render` inside the workspace, producing a `.docx` file styled by the reference template. |
| 6 | Patch headers | Opens the rendered `.docx` and replaces the `Head-Title` and `Head-Subtitle` placeholder strings in every page header. Also marks the TOC field as dirty so Word refreshes it on next open. |
| 7 | Convert to PDF | Converts the patched `.docx` to a `.pdf` alongside it. |
| 8 | Collect output | Copies the final `.docx` and `.pdf` to the directory where you ran the command. |
| 9 | Clean up | Deletes the UUID workspace folder. The `render/` parent directory remains but is otherwise empty. |

---

## Requirements

Before using qlon, make sure the following are installed on your system:

### Python 3.13+

Download from [python.org](https://www.python.org/downloads/). Verify with:

```bash
python --version
```

### Quarto 1.x

Download from [quarto.org](https://quarto.org/docs/get-started/). Verify with:

```bash
quarto --version
```

Both must be available on your system `PATH`. The executables will check for both and print a clear message if either is missing.

---

## Installation

**1. Clone the repository**

```bash
git clone <repo-url>
cd <project-folder>
```

**2. Create a virtual environment**

Using a virtual environment is strongly recommended to isolate dependencies from your system Python.

```bash
python -m venv .venv
```

Supported virtual environment names that qlon detects automatically: `.venv`, `venv`, `env`. Place it at the project root and the executable will use it without any extra configuration.

**3. Activate the environment and install dependencies**

Windows:
```bat
.venv\Scripts\activate
pip install -r requirements.txt
```

Linux / macOS:
```bash
source .venv/bin/activate
pip install -r requirements.txt
```

> If you use `uv`, `poetry`, `conda`, or any other package manager, install from `requirements.txt` using your preferred tool.

---

## Usage

### Windows

```bat
bin\qlon.bat <config.yml>
```

### Linux / macOS

```bash
chmod +x bin/qlon.sh    # first time only
bin/qlon.sh <config.yml>
```

The output `.docx` and `.pdf` are written to whichever directory you run the command from.

### Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `config` | yes* | Path to the YAML config file that describes your document |
| `--test` | — | Run using the built-in `test/example.yml` instead of a config file |
| `--preset <name>` | — | Use a template from the `template/` folder by name (e.g. `basic`) |
| `--custom <path>` | — | Use any `.docx` file on your machine as the reference template |

*Not required when `--test` is used. `--preset` and `--custom` are mutually exclusive.

### Using a Different Template

By default, Qlon uses the built-in `basic.docx` template. You have two ways to override it:

**`--preset <name>`** — use a template from the built-in `template/` folder by name:

```bat
bin\qlon.bat <config.yml> --preset basic
```

**`--custom <path>`** — use any `.docx` file on your machine:

```bat
bin\qlon.bat <config.yml> --custom path\to\custom.docx
```

The selected template is copied into the render workspace for that run. Your original file is never modified. `--preset` and `--custom` are mutually exclusive.

### Test Run

To verify the installation with the built-in example config:

```bat
bin\qlon.bat --test
```

This runs against `test/example.yml` and uses the sample chapter in `test/`. Output files appear in the current directory.

---

## Config File

Each document you produce is driven by a YAML config file. You can name it anything. Start from `test/example.yml` as a template.

```yaml
cover:
  title: "My Document"
  subtitle: "An Internal Reference"
  author: "Jane Doe"
  date: "1 January 2026"

header:
  title: "My Document"
  subtitle: "An Internal Reference"

content:
  folder: "docs/"
  table:
    title: "Table of Contents"

quarto:
  format:
    docx:
      number-sections: true
```

### Reference

| Key | Required | Description |
|-----|----------|-------------|
| `cover.title` | yes | Main title printed on the cover page |
| `cover.subtitle` | no | Subtitle printed below the title on the cover |
| `cover.author` | no | Author name on the cover page |
| `cover.date` | no | Publication date on the cover page |
| `header.title` | no | Replaces the `Head-Title` placeholder in every page header |
| `header.subtitle` | no | Replaces the `Head-Subtitle` placeholder in every page header |
| `content.folder` | yes | Path to your `.qmd` chapter files, relative to the config file |
| `content.table.title` | no | Title shown on the table of contents page. Defaults to the value set in `component/index.qmd` |
| `quarto` | no | Any Quarto format properties to override. Deep-merged into `config/_quarto.yml` after all other settings are applied |

### Notes

- `content.folder` is resolved relative to the config file, not your current directory. This means you can run `qlon` from anywhere.
- The `quarto:` block supports any valid Quarto format property. However, `toc` is always forced to `false` because the TOC is managed by the cover page (`index.qmd`) — this cannot be overridden.
- Chapter files inside `content.folder` are sorted alphabetically. Name them with a numeric prefix (e.g. `01-intro.qmd`, `02-setup.qmd`) to control order.

---

## Output Structure

Every rendered document follows this fixed order:

1. **Cover page** — generated from `component/index.qmd` using the metadata in your config (`title`, `subtitle`, `author`, `date`)
2. **Table of contents** — also part of the cover page, listing all chapters with page numbers
3. **Content** — your `.qmd` chapter files, in alphabetical order

---

## Customizing the Template

The visual appearance of the output — fonts, heading styles, spacing, colors, header and footer layout — is controlled entirely by the Word reference template (`basic.docx`).

To create your own template, copy `template/basic.docx` and start modifying it in Microsoft Word. Save it anywhere, then point Qlon to it using `--preset` (if placed in the `template/` folder) or `--custom` (for any path on your machine).

### Adding static content to the template

For anything that appears on every page and does not change per document — a company logo, a background image, a fixed label, a decorative element — add it directly inside the **header or footer** of the template using Word's header/footer editing mode.

For content that *does* change per document (like the document title shown in the header), use placeholder text instead. Qlon will replace it at render time. The built-in placeholders are:

| Placeholder | Replaced with |
|-------------|--------------|
| `Head-Title` | `header.title` from your config |
| `Head-Subtitle` | `header.subtitle` from your config |

Place these strings anywhere inside the header or footer of your template — in a text box, a table cell, or a plain paragraph — and Qlon will substitute them automatically. Anything else in the header or footer is left untouched.

The same mechanism applies to the **cover page**. Word supports a separate first page header and footer — enable "Different First Page" in the template, then place any static content (a logo, a background, a decorative element) there. It will appear only on the cover and won't affect the rest of the document.

---

## Writing Content

Chapter files are standard `.qmd` files (Quarto Markdown). Each file becomes one chapter in the final document.

```
docs/
├── 01-introduction.qmd
├── 02-installation.qmd
└── 03-usage.qmd
```

Each file should include a YAML frontmatter block at the top:

```yaml
---
title: "Introduction"
---

Your content here...
```

The Word template (`basic.docx`) is automatically applied to every chapter — no frontmatter configuration needed. To use a different template, pass `--preset` or `--custom` on the command line (see [Usage](#usage)).

---

## Project Structure

```
/
├── bin/
│   ├── qlon.bat              # Windows entry point
│   └── qlon.sh               # Linux / macOS entry point
│
├── component/
│   └── index.qmd             # Cover page and table of contents (always the first chapter)
│
├── config/
│   └── _quarto.yml           # Base Quarto project configuration
│
├── script/
│   └── main.py               # Core render pipeline (all logic lives here)
│
├── template/
│   └── basic.docx            # Word reference template (controls fonts, styles, headers)
│
├── test/
│   ├── content.qmd           # Sample chapter file used by --test
│   └── example.yml           # Annotated example config — start here
│
├── render/                   # Parent folder for temporary workspaces — never committed
│   └── <uuid>/               # Per-run isolated workspace, deleted automatically after each run
│
├── requirements.txt          # Python dependencies
└── README.md
```

### Key Files Explained

| File | Purpose |
|------|---------|
| `bin/qlon.bat` / `qlon.sh` | Entry points. Detect virtual environment, check Python and Quarto availability, then launch `script/main.py` |
| `script/main.py` | The full render pipeline — all 8 steps in one file |
| `config/_quarto.yml` | The base Quarto config. Modified at render time — never edited directly |
| `component/index.qmd` | The cover and TOC page. Always included as the first chapter |
| `template/basic.docx` | The Word style reference. Controls all visual formatting in the output |
| `test/example.yml` | A fully annotated config file — use as a starting point for new documents |

---

## Troubleshooting

### `Python is not installed or not found in PATH`

**Cause:** The executable could not find a Python interpreter — neither in a local virtual environment nor on the system PATH.

**Solution:**
- Install Python 3.13+ from [python.org](https://www.python.org/downloads/)
- Ensure `python` is accessible on your PATH by running `python --version` in a terminal
- Alternatively, place a virtual environment named `.venv`, `venv`, or `env` at the project root — Qlon will detect and use it automatically

---

### `Quarto is not installed or not found in PATH`

**Cause:** The `quarto` command could not be found on the system PATH.

**Solution:**
- Install Quarto from [quarto.org](https://quarto.org/docs/get-started/)
- After installation, verify with `quarto --version` in a new terminal
- On Windows, restart your terminal after installation so the PATH change takes effect

---

### `Some required packages are not installed`

**Cause:** One or more Python dependencies (`pyyaml`, `python-docx`, `docx2pdf`, `rich`) are missing from the active Python environment.

**Solution:**
- Activate your virtual environment first, then run:
  ```bash
  pip install -r requirements.txt
  ```
- If you use a different package manager (uv, poetry, conda), install from `requirements.txt` using your tool of choice

---

### `Quarto render failed`

**Cause:** Quarto encountered an error during rendering. The full error output from Quarto is printed to the terminal.

**Common causes and fixes:**

| Cause | Fix |
|-------|-----|
| `reference-doc: template.docx` not found | The template is staged automatically — this usually means a previous run crashed mid-way and left a stale UUID folder in `render/`. Delete any subfolders inside `render/` manually and try again |
| Invalid `.qmd` syntax | Check the Quarto error output for the file and line number, then fix the syntax in your source file |
| Quarto version mismatch | Run `quarto --version` and compare against the version used when the project was set up. Upgrade or downgrade as needed |
| Missing Quarto extension or format | Ensure your Quarto installation supports the `docx` output format. Run `quarto check` for a full environment report |

---

### Output files have incorrect or missing header text

**Cause:** The `Head-Title` or `Head-Subtitle` placeholder strings in the Word template were not found or did not match.

**Solution:**
- Open `template/basic.docx` and verify that the header text contains the exact strings `Head-Title` and `Head-Subtitle` as placeholders
- Check that `header.title` and `header.subtitle` are set correctly in your config file
- Placeholders are case-sensitive — `head-title` will not match `Head-Title`

---

### The table of contents is not updating in Word

**Cause:** Word does not refresh the TOC automatically when a document is first opened.

**Solution:** Open the output `.docx` in Microsoft Word and press `Ctrl+A` then `F9` to update all fields, including the TOC. This is expected behaviour — the TOC field is marked as dirty by Qlon so Word knows to refresh it on the next open.
