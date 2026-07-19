"""Pure helpers for the web layer: preset discovery, config assembly, and packaging
the render output for download. No Flask or request state here.
"""

import io
import zipfile
from pathlib import Path

from paths import TEMPLATE_DIR

DOCX_MIMETYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def list_presets() -> list[str]:
    """Return built-in template names (without the .docx suffix)."""
    if not TEMPLATE_DIR.is_dir():
        return []
    return sorted(p.stem for p in TEMPLATE_DIR.glob("*.docx"))


def build_config(form, content_folder: str) -> dict:
    """Assemble a Qlon config dict from submitted form fields."""
    cover = {"title": form.get("cover_title", "").strip() or "Untitled"}
    for key, field in (("subtitle", "cover_subtitle"), ("author", "cover_author"), ("date", "cover_date")):
        val = form.get(field, "").strip()
        if val:
            cover[key] = val

    header = {}
    for key, field in (("title", "header_title"), ("subtitle", "header_subtitle")):
        val = form.get(field, "").strip()
        if val:
            header[key] = val

    content = {"folder": content_folder}
    table_title = form.get("table_title", "").strip()
    if table_title:
        content["table"] = {"title": table_title}

    config = {"cover": cover, "content": content}
    if header:
        config["header"] = header
    return config


def package_output(docx_path: Path, image_dir: Path) -> tuple[bytes, str, str]:
    """Read the render output into memory, ready to send.

    Returns ``(data, mimetype, download_name)``. When the render produced diagrams
    (a non-empty ``image_dir``), the DOCX is zipped together with the ``Image/`` folder;
    otherwise the bare DOCX bytes are returned. Reading into memory lets the caller
    delete the job workspace immediately after.
    """
    has_images = image_dir.is_dir() and any(image_dir.iterdir())
    if has_images:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(docx_path, docx_path.name)
            for img in sorted(image_dir.iterdir()):
                zf.write(img, f"Image/{img.name}")
        return buf.getvalue(), "application/zip", f"{docx_path.stem}.zip"

    return docx_path.read_bytes(), DOCX_MIMETYPE, docx_path.name


def package_folder(root: Path, zip_stem: str) -> tuple[bytes, str, str]:
    """Zip every file under *root* into memory, ready to send.

    Arc names are kept relative to *root* (so the archive mirrors the folder tree),
    and the download is named ``{zip_stem}.zip``. Reading into memory lets the caller
    delete the job workspace immediately after. Used for the reverse pipeline output,
    which is a whole page/media folder rather than a single file.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(p for p in root.rglob("*") if p.is_file()):
            zf.write(path, path.relative_to(root).as_posix())
    return buf.getvalue(), "application/zip", f"{zip_stem}.zip"
