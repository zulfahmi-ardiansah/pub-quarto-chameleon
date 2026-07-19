"""Flask web interface for Qlon.

Wraps the existing CLI render pipeline (script/main.py). For each request it builds
an isolated job workspace, assembles a config.yml from form fields, stages the
uploaded markdown files, shells out to the CLI, then streams the resulting DOCX
(zipped with the Image/ folder when diagrams are present) back to the browser.

The pipeline is invoked as a subprocess rather than imported because main.py relies
on module-level globals, the current working directory, and rich console output --
shelling out reuses the tested code path without untangling any of that.
"""

import io
import shutil
import subprocess
import sys
import uuid
import zipfile
from pathlib import Path

import yaml
from flask import Flask, abort, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from zip_intake import ZipIntakeError, stage_zip

ROOT_DIR = Path(__file__).resolve().parent.parent
MAIN_SCRIPT = ROOT_DIR / "script" / "main.py"
TEMPLATE_DIR = ROOT_DIR / "template"
JOBS_DIR = Path(__file__).resolve().parent / "jobs"

ALLOWED_CONTENT = {".md", ".qmd"}
RENDER_TIMEOUT = 600  # seconds

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB upload cap
# Pick up template edits without restarting the server (handy with mounted volumes).
app.config["TEMPLATES_AUTO_RELOAD"] = True

STATIC_DIR = Path(__file__).resolve().parent / "static"


@app.context_processor
def inject_asset_version():
    """Cache-busting token = newest static-file mtime. Changes whenever a CSS/JS/asset
    is edited, so templates emit a fresh ?v=… and the browser refetches (no stale cache),
    while unchanged assets still cache normally."""
    try:
        latest = max((f.stat().st_mtime for f in STATIC_DIR.rglob("*") if f.is_file()), default=0)
    except OSError:
        latest = 0
    return {"asset_v": int(latest)}


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


@app.get("/")
def index():
    return render_template("index.html", presets=list_presets())


@app.get("/template/download/<name>")
def download_template(name: str):
    """Serve a built-in template .docx for the user to download and customise."""
    # Sanitise: only allow names that exist as actual files (no path traversal).
    safe = Path(name).name
    path = TEMPLATE_DIR / f"{safe}.docx"
    if not path.is_file():
        abort(404)
    return send_file(path, as_attachment=True, download_name=f"{safe}.docx")


@app.post("/render")
def render():
    files = [f for f in request.files.getlist("files") if f and f.filename]
    pasted = request.form.get("markdown", "").strip()
    if not files and not pasted:
        return jsonify(error="Provide markdown text or upload at least one .md/.qmd file."), 400

    job_dir = JOBS_DIR / uuid.uuid4().hex
    content_dir = job_dir / "content"
    content_dir.mkdir(parents=True, exist_ok=True)

    try:
        # Stage content. Uploaded files take precedence; otherwise the pasted markdown
        # becomes a single chapter. Filename order = chapter order (sorted alphabetically).
        # A single .zip is a project archive: chapter md + referenced images unpacked
        # via stage_zip. Zip and loose md are mutually exclusive.
        zips = [f for f in files if Path(f.filename).suffix.lower() == ".zip"]
        if zips:
            if len(files) > 1:
                return jsonify(error="Upload a project .zip on its own, not mixed with other files."), 400
            archive = job_dir / secure_filename(zips[0].filename)
            zips[0].save(archive)
            try:
                stage_zip(archive, content_dir)
            except ZipIntakeError as exc:
                return jsonify(error=str(exc)), 400
        elif files:
            for f in files:
                name = secure_filename(f.filename)
                if Path(name).suffix.lower() not in ALLOWED_CONTENT:
                    return jsonify(error=f"Unsupported file type: {f.filename}"), 400
                f.save(content_dir / name)
        else:
            (content_dir / "document.md").write_text(pasted, encoding="utf-8")

        # Resolve template selection: custom upload wins over preset.
        cli_args: list[str] = []
        custom_upload = request.files.get("custom_template")
        preset = request.form.get("preset", "").strip()
        if custom_upload and custom_upload.filename:
            if Path(custom_upload.filename).suffix.lower() != ".docx":
                return jsonify(error="Custom template must be a .docx file."), 400
            tpl_dir = job_dir / "_tpl"
            tpl_dir.mkdir(exist_ok=True)
            tpl_path = tpl_dir / secure_filename(custom_upload.filename)
            custom_upload.save(tpl_path)
            cli_args = ["--custom", str(tpl_path)]
        elif preset:
            if preset not in list_presets():
                return jsonify(error=f"Unknown template preset: {preset}"), 400
            cli_args = ["--preset", preset]

        # Write config.yml at the job root; content.folder resolves relative to it.
        config = build_config(request.form, "content/")
        config_path = job_dir / "config.yml"
        config_path.write_text(yaml.safe_dump(config, sort_keys=False, allow_unicode=True), encoding="utf-8")

        # Run the existing CLI pipeline with the job dir as the working directory so
        # outputs (DOCX + Image/) land inside the job dir.
        proc = subprocess.run(
            [sys.executable, str(MAIN_SCRIPT), "config.yml", *cli_args],
            cwd=job_dir,
            capture_output=True,
            text=True,
            timeout=RENDER_TIMEOUT,
        )
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "Render failed with no output.").strip()
            return jsonify(error="Render failed.", detail=detail[-4000:]), 500

        docx_files = sorted(p for p in job_dir.glob("*.docx"))
        if not docx_files:
            return jsonify(error="Render finished but no DOCX was produced.", detail=(proc.stdout or "")[-4000:]), 500
        docx_path = docx_files[0]

        image_dir = job_dir / "Image"
        has_images = image_dir.is_dir() and any(image_dir.iterdir())

        # Read payload into memory before the job dir is cleaned up in `finally`.
        if has_images:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.write(docx_path, docx_path.name)
                for img in sorted(image_dir.iterdir()):
                    zf.write(img, f"Image/{img.name}")
            buf.seek(0)
            return send_file(buf, mimetype="application/zip", as_attachment=True,
                             download_name=f"{docx_path.stem}.zip")

        buf = io.BytesIO(docx_path.read_bytes())
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=docx_path.name,
        )
    except subprocess.TimeoutExpired:
        return jsonify(error=f"Render timed out after {RENDER_TIMEOUT}s."), 504
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
