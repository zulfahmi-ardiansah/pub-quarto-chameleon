"""Flask web interface for Qlon — HTTP controllers only.

Routes parse the request, delegate to the service/util layers, and translate the result
(or a :class:`RenderError`) into an HTTP response. The render orchestration lives in
``render.py``, reverse in ``reverse.py``, content staging in ``staging.py``, and pure
helpers in ``util.py``.
"""

import io
from pathlib import Path

from flask import Flask, abort, jsonify, render_template, request, send_file

from paths import STATIC_DIR, TEMPLATE_DIR
from render import RenderError, render_document
from reverse import reverse_document
from util import list_presets

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB upload cap
# Pick up template edits without restarting the server (handy with mounted volumes).
app.config["TEMPLATES_AUTO_RELOAD"] = True


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
    try:
        result = render_document(
            files, request.files.get("custom_template"), request.form, pasted)
    except RenderError as exc:
        payload = {"error": exc.message}
        if exc.detail:
            payload["detail"] = exc.detail
        return jsonify(payload), exc.status

    return send_file(
        io.BytesIO(result.data),
        mimetype=result.mimetype,
        as_attachment=True,
        download_name=result.download_name,
    )


@app.post("/reverse")
def reverse():
    """Convert an uploaded .docx into a zip of clean Markdown pages (+ media)."""
    try:
        result = reverse_document(request.files.getlist("docx"), request.form)
    except RenderError as exc:
        payload = {"error": exc.message}
        if exc.detail:
            payload["detail"] = exc.detail
        return jsonify(payload), exc.status

    return send_file(
        io.BytesIO(result.data),
        mimetype=result.mimetype,
        as_attachment=True,
        download_name=result.download_name,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
