#!/usr/bin/env bash
ROOT="$(dirname "$0")/.."

for venv in ".venv" "venv" "env"; do
    if [ -f "$ROOT/$venv/bin/python" ]; then
        exec "$ROOT/$venv/bin/python" "$ROOT/script/main.py" "$@"
    fi
done

if ! command -v python &>/dev/null; then
    echo "Python is not installed or not found in PATH."
    exit 1
fi

if ! command -v quarto &>/dev/null; then
    echo "Quarto is not installed or not found in PATH."
    exit 1
fi

exec python "$ROOT/script/main.py" "$@"
