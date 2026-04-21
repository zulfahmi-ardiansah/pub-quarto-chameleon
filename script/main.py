"""Render a Quarto book to DOCX and PDF from a user-supplied YAML config.

Usage:
    qlon <config.yml>

The script stages all assets into the render/ workspace, runs Quarto, post-processes
the output (header replacements, TOC refresh, PDF conversion), copies results to the
caller's working directory, then cleans up the workspace.
"""

import argparse
import contextlib
import importlib.util
import io
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

# Check third-party dependencies before importing them so the error is actionable
_REQUIRED = {"yaml": "pyyaml", "docx": "python-docx", "docx2pdf": "docx2pdf", "rich": "rich"}
_missing = [pip for mod, pip in _REQUIRED.items() if importlib.util.find_spec(mod) is None]
if _missing:
    print("Some required packages are not installed:")
    for pkg in _missing:
        print(f"    - {pkg}")
    print("Please install all dependencies from requirements.txt")
    sys.exit(1)

import yaml
from docx import Document
from docx.oxml.ns import qn
from docx2pdf import convert as docx2pdf_convert
from rich.console import Console
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn

ROOT_DIR = Path(__file__).parent.parent
CONFIG_DIR = ROOT_DIR / "config"
COMPONENT_DIR = ROOT_DIR / "component"
TEMPLATE_DIR = ROOT_DIR / "template"
RENDER_BASE_DIR = ROOT_DIR / "render"
RENDER_DIR = RENDER_BASE_DIR  # reassigned to a UUID subfolder at runtime
OUTPUT_DIR = RENDER_DIR / "_output"  # follows RENDER_DIR at runtime

_TOC_INSTR = "TOC"
_PLACEHOLDER_TITLE = "Head-Title"
_PLACEHOLDER_SUBTITLE = "Head-Subtitle"
_console = Console()


def deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge *override* into *base* in place and return *base*."""
    for key, val in override.items():
        if key in base and isinstance(base[key], dict) and isinstance(val, dict):
            deep_merge(base[key], val)
        else:
            base[key] = val
    return base


def setup_workspace(custom_template: Path | None = None) -> None:
    """Populate the UUID workspace inside render/ with the base Quarto config, cover page, and Word template."""
    RENDER_BASE_DIR.mkdir(exist_ok=True)
    RENDER_DIR.mkdir()
    shutil.copy(CONFIG_DIR / "_quarto.yml", RENDER_DIR / "_quarto.yml")
    shutil.copy(COMPONENT_DIR / "index.qmd", RENDER_DIR / "index.qmd")
    template_src = custom_template if custom_template else TEMPLATE_DIR / "basic.docx"
    shutil.copy(template_src, RENDER_DIR / "template.docx")


def patch_index_title(title: str) -> None:
    """Replace the title in the render/index.qmd frontmatter."""
    index_qmd = RENDER_DIR / "index.qmd"
    text = index_qmd.read_text(encoding="utf-8")
    parts = text.split("---", 2)
    frontmatter = yaml.safe_load(parts[1])
    frontmatter["title"] = title
    parts[1] = "\n" + yaml.dump(frontmatter, allow_unicode=True, default_flow_style=False)
    index_qmd.write_text("---".join(parts), encoding="utf-8")


def copy_content(content_folder: str, config_dir: Path) -> list[str]:
    """Copy all .qmd files from *content_folder* into render/ and return their names."""
    src = (config_dir / content_folder).resolve()
    names = []
    for f in sorted(src.glob("*.qmd")):
        shutil.copy(f, RENDER_DIR / f.name)
        names.append(f.name)
    return names


def configure_quarto(config: dict, chapter_files: list[str]) -> None:
    """Patch render/_quarto.yml with cover metadata, chapter list, and quarto format overrides."""
    quarto_yml = RENDER_DIR / "_quarto.yml"
    data = yaml.safe_load(quarto_yml.read_text(encoding="utf-8"))

    cover = config.get("cover", {})
    book = data.setdefault("book", {})
    book.update({k: cover[k] for k in ("title", "date", "author", "subtitle") if k in cover})
    book["chapters"] = ["index.qmd"] + chapter_files

    quarto_overrides = config.get("quarto", {})
    if quarto_overrides:
        deep_merge(data, quarto_overrides)

    # toc is managed by the cover page (index.qmd); user overrides are not allowed
    data.setdefault("format", {}).setdefault("docx", {})["toc"] = False

    quarto_yml.write_text(yaml.dump(data, allow_unicode=True, default_flow_style=False), encoding="utf-8")


def quarto_render() -> None:
    """Run `quarto render` inside render/ and exit on failure."""
    result = subprocess.run(["quarto", "render"], capture_output=True, cwd=RENDER_DIR)
    if result.returncode != 0:
        _console.print("[red]Quarto render failed.[/red]")
        _console.print(result.stderr.decode())
        sys.exit(result.returncode)


def replace_in_header(doc: Document, replacements: dict[str, str]) -> None:
    """Replace placeholder strings in all section headers of *doc*.

    Both w:t (body text runs) and a:t (DrawingML text) are scanned because
    Word sometimes places header content inside drawing objects.
    """
    tags = (qn("w:t"), qn("a:t"))
    for section in doc.sections:
        for header in (section.header, section.even_page_header, section.first_page_header):
            for t_elem in header._element.iter(*tags):
                if not t_elem.text:
                    continue
                for search, replace in replacements.items():
                    if search in t_elem.text:
                        t_elem.text = t_elem.text.replace(search, replace)


def mark_toc_dirty(doc: Document) -> None:
    """Set w:dirty="1" on TOC field starts so Word refreshes the TOC on next open."""
    for fld_char in doc.element.body.iter(qn("w:fldChar")):
        if fld_char.get(qn("w:fldCharType")) != "begin":
            continue
        for sibling in fld_char.getparent().itersiblings():
            instr = next(sibling.iter(qn("w:instrText")), None)
            if instr is not None:
                if instr.text and _TOC_INSTR in instr.text:
                    fld_char.set(qn("w:dirty"), "1")
                break


def patch_headers(replacements: dict[str, str]) -> None:
    """Apply header replacements and TOC dirty-flag to every DOCX in the output dir."""
    for docx_file in OUTPUT_DIR.glob("*.docx"):
        doc = Document(docx_file)
        replace_in_header(doc, replacements)
        mark_toc_dirty(doc)
        doc.save(docx_file)


def convert_to_pdf() -> None:
    """Convert every DOCX in the output dir to PDF alongside the original file."""
    for docx_file in OUTPUT_DIR.glob("*.docx"):
        # Suppress docx2pdf's own progress output to keep the rich display clean
        with contextlib.redirect_stderr(io.StringIO()):
            docx2pdf_convert(docx_file, docx_file.with_suffix(".pdf"))


def collect_output(dest: Path) -> None:
    """Copy all DOCX and PDF results from the output dir to *dest*."""
    for f in OUTPUT_DIR.iterdir():
        if f.suffix in (".docx", ".pdf"):
            shutil.copy(f, dest / f.name)


def cleanup() -> None:
    """Delete the UUID workspace folder inside render/."""
    shutil.rmtree(RENDER_DIR)


def main() -> None:
    """Entry point: parse args, orchestrate the 8-step render pipeline."""
    parser = argparse.ArgumentParser(description="Render a Quarto book to DOCX and PDF.")
    parser.add_argument("config", nargs="?", help="Path to input yml file")
    parser.add_argument("--test", action="store_true", help="Run using example.yml")
    parser.add_argument("--preset", metavar="NAME", help="Name of a built-in template in the template/ folder (e.g. basic)")
    parser.add_argument("--custom", metavar="PATH", help="Path to a custom .docx reference template")
    args = parser.parse_args()

    if args.test:
        config_path = ROOT_DIR / "test" / "example.yml"
    elif args.config:
        config_path = Path(args.config)
    else:
        parser.error("provide a config file or use --test")

    try:
        config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        _console.print(f"[red]File not found: {config_path}[/red]")
        sys.exit(1)

    global RENDER_DIR, OUTPUT_DIR
    RENDER_DIR = RENDER_BASE_DIR / str(uuid.uuid4())
    OUTPUT_DIR = RENDER_DIR / "_output"

    config_dir = config_path.resolve().parent
    original_cwd = Path.cwd()
    if args.preset and args.custom:
        _console.print("[red]Use either --preset or --custom, not both.[/red]")
        sys.exit(1)

    if args.preset:
        custom_template = TEMPLATE_DIR / f"{args.preset}.docx"
        if not custom_template.exists():
            _console.print(f"[red]Preset not found: {custom_template}[/red]")
            sys.exit(1)
    elif args.custom:
        custom_template = Path(args.custom).resolve()
        if not custom_template.exists():
            _console.print(f"[red]Template not found: {custom_template}[/red]")
            sys.exit(1)
    else:
        custom_template = None
    header = config.get("header", {})
    content = config.get("content", {})
    table_title = content.get("table", {}).get("title")

    progress_columns = (
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TaskProgressColumn(),
    )

    _console.print("Run Qlon")

    with Progress(*progress_columns, console=_console) as progress:
        total_steps = 8 + (1 if table_title else 0)
        task = progress.add_task("Setting up workspace...", total=total_steps)

        setup_workspace(custom_template)
        progress.advance(task)

        if table_title:
            progress.update(task, description="Patching table of contents title...")
            patch_index_title(table_title)
            progress.advance(task)

        progress.update(task, description="Copying content files...")
        chapter_files = copy_content(content.get("folder", "content/"), config_dir)
        progress.advance(task)

        progress.update(task, description="Configuring Quarto...")
        configure_quarto(config, chapter_files)
        progress.advance(task)

        progress.update(task, description="Rendering with Quarto...")
        quarto_render()
        progress.advance(task)

        progress.update(task, description="Patching headers...")
        patch_headers({
            _PLACEHOLDER_TITLE: header.get("title", ""),
            _PLACEHOLDER_SUBTITLE: header.get("subtitle", ""),
        })
        progress.advance(task)

        progress.update(task, description="Converting to PDF...")
        convert_to_pdf()
        progress.advance(task)

        progress.update(task, description="Collecting output...")
        collect_output(original_cwd)
        progress.advance(task)

        progress.update(task, description="Cleaning up...")
        cleanup()
        progress.advance(task)

    _console.print("[green]Done![/green]")


if __name__ == "__main__":
    main()
