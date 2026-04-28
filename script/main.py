"""Render a Quarto book to DOCX from a user-supplied YAML config.

Usage:
    qlon <config.yml>

The script stages all assets into the render/ workspace, runs Quarto, post-processes
the output (header replacements, TOC refresh), copies results to the caller's working
directory, then cleans up the workspace.
"""

import argparse
import importlib.util
import re
import shutil
import subprocess
import sys
import uuid
from pathlib import Path

# Check third-party dependencies before importing them so the error is actionable
_REQUIRED = {"yaml": "pyyaml", "docx": "python-docx", "rich": "rich", "playwright": "playwright"}
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
from playwright.sync_api import sync_playwright
from rich.console import Console
from rich.panel import Panel
from rich.progress import BarColumn, Progress, SpinnerColumn, TaskProgressColumn, TextColumn
from rich.text import Text

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
_IMG_REF = re.compile(r'!\[.*?\]\(([^)]+)\)')
_MERMAID_BLOCK = re.compile(r'```mermaid\n(.*?)```', re.DOTALL)
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


def _playwright_render_mermaid(browser, diagram_src: str) -> bytes:
    import html as _html
    context = browser.new_context(device_scale_factor=3)
    page = context.new_page()
    page.set_content(
        "<!DOCTYPE html><html><head>"
        '<script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>'
        "</head><body style='margin:0;background:white;'>"
        f"<div class='mermaid' style='display:inline-block;padding:16px;'>{_html.escape(diagram_src)}</div>"
        "<script>mermaid.initialize({startOnLoad:true});</script>"
        "</body></html>"
    )
    page.wait_for_selector(".mermaid svg", timeout=15_000)
    png = page.locator(".mermaid").screenshot()
    context.close()
    return png


def _render_mermaid_blocks(text: str, dest_dir: Path, names: list[str]) -> tuple[str, dict[str, Path]]:
    """Render each ```mermaid block to a PNG and replace it with an image reference.

    *names* must have one entry per mermaid block — used as the output filename.
    Returns (modified_text, {ref_string: created_path}) for every diagram rendered.
    Falls back to Quarto-native ```{mermaid} syntax on failure (not included in the dict).
    """
    matches = list(_MERMAID_BLOCK.finditer(text))
    if not matches:
        return text, {}

    diagrams_dir = dest_dir / "diagrams"
    diagrams_dir.mkdir(parents=True, exist_ok=True)

    replacements: dict[int, str] = {}
    created: dict[str, Path] = {}
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        for match, name in zip(matches, names):
            try:
                png = _playwright_render_mermaid(browser, match.group(1))
                path = diagrams_dir / name
                path.write_bytes(png)
                ref = f"diagrams/{name}"
                replacements[match.start()] = f"![]({ref})"
                created[ref] = path
            except Exception:
                replacements[match.start()] = f"```{{mermaid}}\n{match.group(1)}```"
        browser.close()

    result = text
    for match in reversed(matches):
        result = result[:match.start()] + replacements[match.start()] + result[match.end():]
    return result, created


def _copy_images(text: str, src_dir: Path) -> None:
    """Copy relative images referenced in *text* from *src_dir* into the workspace."""
    for match in _IMG_REF.finditer(text):
        raw = match.group(1).split()[0]  # strip optional title attribute
        if raw.startswith(("http://", "https://", "/")):
            continue
        img_src = (src_dir / raw).resolve()
        if img_src.exists():
            img_dest = RENDER_DIR / raw
            img_dest.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(img_src, img_dest)


def copy_content(content_folder: str, config_dir: Path) -> tuple[list[str], list[list[Path]]]:
    """Copy all .qmd and .md files from *content_folder* into the workspace.

    .md files are copied as .qmd — Quarto picks up their H1 as the chapter title natively.
    Relative images referenced in any file are also copied into the workspace.
    Returns the list of .qmd filenames and a per-chapter ordered list of image paths.
    """
    src = (config_dir / content_folder).resolve()
    files = sorted(list(src.glob("*.qmd")) + list(src.glob("*.md")), key=lambda f: f.name)
    names: list[str] = []
    chapter_images: list[list[Path]] = []
    for chapter_idx, f in enumerate(files, 1):
        text = f.read_text(encoding="utf-8")
        _copy_images(text, f.parent)

        # Assign y-indices by merging user image refs and mermaid blocks in document order
        events: list[tuple[int, str]] = []
        seen_refs: set[str] = set()
        for m in _IMG_REF.finditer(text):
            raw = m.group(1).split()[0]
            if raw.startswith(("http://", "https://", "/")) or raw in seen_refs:
                continue
            seen_refs.add(raw)
            events.append((m.start(), "img"))
        for m in _MERMAID_BLOCK.finditer(text):
            events.append((m.start(), "mermaid"))
        events.sort()
        mermaid_names = [
            f"Diagram-{chapter_idx}-{y}.png"
            for y, (_, kind) in enumerate(events, 1)
            if kind == "mermaid"
        ]

        text, diagram_paths = _render_mermaid_blocks(text, RENDER_DIR, mermaid_names)
        dest_name = f.stem + ".qmd" if f.suffix == ".md" else f.name
        (RENDER_DIR / dest_name).write_text(text, encoding="utf-8")
        names.append(dest_name)

        seen: set[str] = set()
        imgs: list[Path] = []
        for match in _IMG_REF.finditer(text):
            raw = match.group(1).split()[0]
            if raw.startswith(("http://", "https://", "/")) or raw in seen:
                continue
            seen.add(raw)
            p = diagram_paths.get(raw) or RENDER_DIR / raw
            if p.exists():
                imgs.append(p)
        chapter_images.append(imgs)
    return names, chapter_images


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


def collect_images(dest: Path, chapter_images: list[list[Path]]) -> None:
    """Copy images into dest/Image/ named Diagram-{chapter}-{index}.{ext}."""
    if not any(chapter_images):
        return
    img_dir = dest / "Image"
    img_dir.mkdir(exist_ok=True)
    for x, images in enumerate(chapter_images, 1):
        for y, f in enumerate(images, 1):
            shutil.copy(f, img_dir / f"Diagram-{x}-{y}{f.suffix.lower()}")


def collect_output(dest: Path) -> None:
    """Copy all DOCX results from the output dir to *dest*."""
    for f in OUTPUT_DIR.iterdir():
        if f.suffix == ".docx":
            shutil.copy(f, dest / f.name)


def cleanup() -> None:
    """Delete the UUID workspace folder inside render/."""
    shutil.rmtree(RENDER_DIR)


def main() -> None:
    """Entry point: parse args, orchestrate the 8-step render pipeline."""
    parser = argparse.ArgumentParser(description="Render a Quarto book to DOCX.")
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

    cover = config.get("cover", {})
    doc_title = cover.get("title", "Untitled")
    doc_subtitle = cover.get("subtitle", "")
    doc_author = cover.get("author", "")

    header_text = Text()
    header_text.append(doc_title, style="bold white")
    if doc_subtitle:
        header_text.append(f"\n{doc_subtitle}", style="dim white")
    if doc_author:
        header_text.append(f"\n{doc_author}", style="italic dim")
    header_text.append(f"\n{config_path.name}", style="dim cyan")

    _console.print(Panel(header_text, title="[bold cyan]Qlon[/bold cyan]", border_style="cyan", padding=(0, 2)))

    with Progress(*progress_columns, console=_console) as progress:
        total_steps = 8 + (1 if table_title else 0)
        task = progress.add_task("Preparing render workspace...", total=total_steps)

        setup_workspace(custom_template)
        progress.advance(task)

        if table_title:
            progress.update(task, description="Setting custom table of contents title...")
            patch_index_title(table_title)
            progress.advance(task)

        progress.update(task, description="Copying and pre-processing content files...")
        chapter_files, chapter_images = copy_content(content.get("folder", "content/"), config_dir)
        progress.advance(task)

        progress.update(task, description="Writing Quarto configuration...")
        configure_quarto(config, chapter_files)
        progress.advance(task)

        progress.update(task, description="Rendering document with Quarto (this may take a while)...")
        quarto_render()
        progress.advance(task)

        progress.update(task, description="Applying header replacements and marking TOC for refresh...")
        patch_headers({
            _PLACEHOLDER_TITLE: header.get("title", ""),
            _PLACEHOLDER_SUBTITLE: header.get("subtitle", ""),
        })
        progress.advance(task)

        progress.update(task, description="Saving DOCX to output directory...")
        collect_output(original_cwd)
        progress.advance(task)

        has_images = any(chapter_images)
        progress.update(task, description="Exporting images to Image/ folder..." if has_images else "No images found, skipping export...")
        collect_images(original_cwd, chapter_images)
        progress.advance(task)

        progress.update(task, description="Cleaning up render workspace...")
        cleanup()
        progress.advance(task)

    _console.print("\n[bold green]Done![/bold green] Your DOCX is ready in the current directory.\n")

    notes = Text()
    notes.append("1. ", style="bold yellow")
    notes.append("Open the DOCX — Word will prompt to update fields. Accept it to refresh the Table of Contents.\n")
    notes.append("2. ", style="bold yellow")
    notes.append("Check every image in the document. ")
    notes.append("Images are not auto-sized — resize them manually in Word as needed.\n", style="bold")
    notes.append("3. ", style="bold yellow")
    notes.append("Mermaid diagrams are pre-rendered as PNG. If a diagram looks too small or blurry, scale it up in Word.\n")
    
    if has_images:
        notes.append("4. ", style="bold yellow")
        notes.append("All images are also saved to ", style="")
        notes.append("./Image/", style="bold cyan")
        notes.append(" for reference or reuse.")

    _console.print(Panel(notes, title="[bold]Post-render checklist[/bold]", border_style="yellow", padding=(1, 2)))


if __name__ == "__main__":
    main()
