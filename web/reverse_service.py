"""Reverse orchestration: turn an uploaded .docx into downloadable Markdown pages.

`reverse_document` owns one reverse (ingest) job end to end — it builds an isolated
workspace, saves the upload, translates the submitted options into CLI flags, shells
out to the same ``script/main.py`` entry point (a ``.docx`` positional routes it to the
reverse pipeline), then zips the produced page/media folder — and tears the workspace
down. Like ``service.py`` it stays free of Flask: plain form data and a file object in,
a :class:`RenderResult` out (or a :class:`RenderError`), leaving HTTP to ``app.py``.
"""

import shutil
import uuid
from pathlib import Path

from werkzeug.utils import secure_filename

from paths import JOBS_DIR
from service import RenderError, RenderResult, run_pipeline
from util import package_folder

REVERSE_TIMEOUT = 900  # seconds — the LLM path can be slow


def _truthy(value: str | None) -> bool:
    """Form checkboxes arrive as 'on'/'true'/'1' (or absent). Coerce to bool."""
    return str(value).strip().lower() in {"on", "true", "1", "yes"}


def build_reverse_args(form, docx_name: str) -> list[str]:
    """Translate the submitted reverse options into the CLI positional + flags.

    *docx_name* is the (already sanitised) filename saved in the job dir; it leads as
    the positional so ``main.py`` routes to the reverse pipeline. Fuma layout implies
    the LLM, so ``--use-llm`` is added whenever fuma is chosen or the toggle is on.
    Model endpoint/name/key and reorder are only meaningful with the LLM.
    """
    layout = (form.get("layout") or "flat").strip().lower()
    if layout not in {"flat", "fuma"}:
        raise RenderError(f"Unknown layout: {layout}")

    use_llm = _truthy(form.get("use_llm")) or layout == "fuma"

    args: list[str] = [docx_name, "--layout", layout]
    if use_llm:
        args.append("--use-llm")
        for flag, field in (("--model-endpoint", "model_endpoint"),
                            ("--model-name", "model_name"),
                            ("--model-key", "model_key")):
            val = (form.get(field) or "").strip()
            if val:
                args += [flag, val]
        if _truthy(form.get("allow_reorder")):
            args.append("--allow-reorder")

    skip_toc = (form.get("skip_toc") or "").strip()
    if skip_toc:
        args += ["--skip-toc", skip_toc]
    if _truthy(form.get("no_merge")):
        args.append("--no-merge")
    return args


def reverse_document(docx_upload, form) -> RenderResult:
    """Run one reverse job and return its packaged Markdown output.

    *docx_upload* is the uploaded ``.docx`` file object, *form* the submitted options.
    Raises :class:`RenderError` (mapped to an HTTP response by the caller) on any
    failure. The whole produced folder (numbered pages + media, or the fuma tree) is
    zipped into memory so the job workspace can be deleted immediately.
    """
    if not (docx_upload and docx_upload.filename):
        raise RenderError("Upload a .docx file to convert.", status=400)
    if Path(docx_upload.filename).suffix.lower() != ".docx":
        raise RenderError("Input must be a .docx file.", status=400)

    filename = secure_filename(docx_upload.filename)
    # secure_filename restricts the charset but keeps a leading '-', which the CLI
    # would parse as a flag rather than the input positional (argv flag smuggling).
    if not filename or filename.startswith("-"):
        raise RenderError("Could not read the uploaded file name.", status=400)
    stem = Path(filename).stem

    job_dir = JOBS_DIR / uuid.uuid4().hex
    job_dir.mkdir(parents=True, exist_ok=True)
    try:
        docx_upload.save(job_dir / filename)

        args = build_reverse_args(form, filename)
        run_pipeline(job_dir, args, timeout=REVERSE_TIMEOUT)

        # The reverse pipeline writes its output into a folder named after the input
        # stem, beside the input (cwd = job_dir). Zip that folder for download.
        out_dir = job_dir / stem
        if not out_dir.is_dir() or not any(out_dir.rglob("*")):
            raise RenderError(
                "Conversion finished but produced no Markdown.", status=500)

        data, mimetype, download_name = package_folder(out_dir, stem)
        return RenderResult(data, mimetype, download_name)
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)
