"""Filesystem locations shared across the web layers."""

from pathlib import Path

WEB_DIR = Path(__file__).resolve().parent
ROOT_DIR = WEB_DIR.parent

MAIN_SCRIPT = ROOT_DIR / "script" / "main.py"   # CLI render/reverse entry point
TEMPLATE_DIR = ROOT_DIR / "template"            # built-in .docx presets
JOBS_DIR = WEB_DIR / "jobs"                      # per-request scratch workspaces
STATIC_DIR = WEB_DIR / "static"                  # CSS/JS/images
