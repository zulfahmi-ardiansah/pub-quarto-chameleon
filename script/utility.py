"""Shared helpers for the qlon pipelines (yml render + docx ingest).

Generic string/dict/env utilities plus a small OpenAI-compatible LLM client
(OpenRouter by default). Imported by render.py and reverse.py.
"""
from __future__ import annotations
import os
import re
import time
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any
import httpx


# ==========================================
# GENERAL UTILITIES
# ==========================================

# --- Static Variable ---

_SLUG_RE = re.compile(r"[^a-z0-9]+")

# --- Utility Method ---

def slugify(text: str) -> str:
    return _SLUG_RE.sub("-", text.lower()).strip("-") or "section"


def deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge *override* into *base* in place and return *base*."""
    for key, val in override.items():
        if key in base and isinstance(base[key], dict) and isinstance(val, dict):
            deep_merge(base[key], val)
        else:
            base[key] = val
    return base


def load_dotenv(path: Path) -> None:
    """Minimal .env loader: KEY=VALUE lines into os.environ (no overwrite)."""
    # is_file (not exists): in this repo `.env` is a virtualenv directory, not a file.
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


# ==========================================
# LLM CLIENT
# ==========================================

# --- Static Variable ---

DEFAULT_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "deepseek/deepseek-v4-flash"
_RETRYABLE_STATUS = frozenset({429, 500, 502, 503, 504})
_DEFAULT_TIMEOUT = httpx.Timeout(120, connect=30)

# Transport/HTTP errors plus the parse/validation errors from a malformed response.
_RETRYABLE_EXC = (httpx.HTTPError, KeyError, IndexError, ValueError)

Message = dict[str, str]

class LLMError(RuntimeError):
    """Raised when a chat completion cannot be obtained after all retries."""


# --- Utility Method ---

def base_url() -> str:
    """API root from ``LLM_BASE_URL``, default OpenRouter, no trailing slash."""
    return os.environ.get("LLM_BASE_URL", DEFAULT_BASE_URL).rstrip("/")


def api_key() -> str | None:
    """Bearer token from ``LLM_API_KEY``, falling back to ``OPENROUTER_API_KEY``."""
    return os.environ.get("LLM_API_KEY") or os.environ.get("OPENROUTER_API_KEY")


def default_model() -> str:
    """Default model id from ``LLM_MODEL`` (or the built-in default)."""
    return os.environ.get("LLM_MODEL", DEFAULT_MODEL)


def _headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        # OpenRouter attribution headers; ignored elsewhere.
        "HTTP-Referer": "https://github.com/local/personal-content-pipeline",
        "X-Title": "personal-content-pipeline",
    }


def chat_completion(
    client: httpx.Client,
    token: str,
    model: str,
    messages: Sequence[Message],
    *,
    temperature: float = 0.2,
    response_format: dict[str, Any] | None = None,
    max_retries: int = 3,
    validate: Callable[[str], None] | None = None,
    on_retry: Callable[[int, Exception], None] | None = None,
) -> str:
    """Send one chat-completion request and return the message content string.

    The endpoint is ``{LLM_BASE_URL}/chat/completions``. Retries retryable HTTP
    statuses and transport/parse errors with exponential backoff (1s, 2s, 4s, ...).
    Raises `LLMError` once `max_retries` is exhausted.

    `validate` runs on the returned content before it is accepted; raising from it
    (e.g. `LLMError` or `ValueError`) rejects the response and triggers a retry, so
    callers can enforce response-shape checks inside the retry loop.
    `on_retry(attempt, error)` is invoked before each backoff sleep for logging.
    """
    url = f"{base_url()}/chat/completions"
    payload: dict[str, Any] = {
        "model": model,
        "messages": list(messages),
        "temperature": temperature,
    }
    if response_format is not None:
        payload["response_format"] = response_format
    headers = _headers(token)

    last_err: Exception | None = None
    for attempt in range(max_retries):
        try:
            resp = client.post(url, json=payload, headers=headers, timeout=_DEFAULT_TIMEOUT)
            if resp.status_code in _RETRYABLE_STATUS:
                raise LLMError(f"retryable status {resp.status_code}: {resp.text[:200]}")
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"].strip()
            if validate is not None:
                validate(content)
            return content
        except (LLMError, *_RETRYABLE_EXC) as e:
            last_err = e
            is_last = attempt == max_retries - 1
            if not is_last:
                if on_retry is not None:
                    on_retry(attempt, e)
                time.sleep(2 ** attempt)
    raise LLMError(f"failed after {max_retries} attempts: {last_err}")
