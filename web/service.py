"""Render orchestration: turn a submitted request into a downloadable document.

`render_document` owns one render job end to end — it builds an isolated workspace,
stages the content, assembles config.yml, shells out to the CLI pipeline, and packages
the output — then tears the workspace down. It stays free of Flask: it takes plain form
data and file objects and returns a :class:`RenderResult` (or raises :class:`RenderError`),
leaving HTTP concerns to the controller in ``app.py``.

The pipeline is invoked as a subprocess rather than imported because ``script/main.py``
relies on module-level globals, the current working directory, and rich console output --
shelling out reuses the tested code path without untangling any of that.
"""

import shutil
import subprocess
import sys
import uuid
from dataclasses import dataclass
from pathlib import Path

import yaml
from werkzeug.utils import secure_filename

from paths import JOBS_DIR, MAIN_SCRIPT
from staging import StagingError, stage_content
from util import build_config, list_presets, package_output

RENDER_TIMEOUT = 600  # seconds


@dataclass
class RenderResult:
    """A finished render, held in memory so the job workspace can be deleted."""
    data: bytes
    mimetype: str
    download_name: str


class RenderError(Exception):
    """A render failure to surface to the client. Carries an HTTP status and detail."""

    def __init__(self, message: str, status: int = 400, detail: str | None = None):
        super().__init__(message)
        self.message = message
        self.status = status
        self.detail = detail


def _resolve_template_args(custom_upload, preset: str, job_dir: Path) -> list[str]:
    """Build the CLI's template flags: a custom .docx upload wins over a named preset."""
    if custom_upload and custom_upload.filename:
        if Path(custom_upload.filename).suffix.lower() != ".docx":
            raise RenderError("Custom template must be a .docx file.")
        tpl_dir = job_dir / "_tpl"
        tpl_dir.mkdir(exist_ok=True)
        tpl_path = tpl_dir / secure_filename(custom_upload.filename)
        custom_upload.save(tpl_path)
        return ["--custom", str(tpl_path)]
    if preset:
        if preset not in list_presets():
            raise RenderError(f"Unknown template preset: {preset}")
        return ["--preset", preset]
    return []


def _run_cli(job_dir: Path, cli_args: list[str]) -> None:
    """Run the CLI render with *job_dir* as cwd, so outputs land inside it."""
    try:
        proc = subprocess.run(
            [sys.executable, str(MAIN_SCRIPT), "config.yml", *cli_args],
            cwd=job_dir,
            capture_output=True,
            text=True,
            timeout=RENDER_TIMEOUT,
        )
    except subprocess.TimeoutExpired:
        raise RenderError(f"Render timed out after {RENDER_TIMEOUT}s.", status=504)
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "Render failed with no output.").strip()
        raise RenderError("Render failed.", status=500, detail=detail[-4000:])


def render_document(files, custom_upload, form, pasted: str) -> RenderResult:
    """Run one render job and return its packaged output.

    *files* are the uploaded content parts, *custom_upload* the optional template file,
    *form* the submitted fields, *pasted* the paste-box Markdown. Raises
    :class:`RenderError` (mapped to an HTTP response by the caller) on any failure.
    """
    job_dir = JOBS_DIR / uuid.uuid4().hex
    content_dir = job_dir / "content"
    content_dir.mkdir(parents=True, exist_ok=True)

    try:
        try:
            stage_content(files, pasted, content_dir)
        except StagingError as exc:
            raise RenderError(str(exc), status=400)

        cli_args = _resolve_template_args(custom_upload, form.get("preset", "").strip(), job_dir)

        # config.yml sits at the job root; content.folder resolves relative to it.
        config = build_config(form, "content/")
        (job_dir / "config.yml").write_text(
            yaml.safe_dump(config, sort_keys=False, allow_unicode=True), encoding="utf-8")

        _run_cli(job_dir, cli_args)

        docx_files = sorted(job_dir.glob("*.docx"))
        if not docx_files:
            raise RenderError("Render finished but no DOCX was produced.", status=500)

        data, mimetype, download_name = package_output(docx_files[0], job_dir / "Image")
        return RenderResult(data, mimetype, download_name)
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)
