#!/usr/bin/env sh
set -eu

PROJECT_ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
PYTHON_COMMAND=${PYTHON:-python3}

"$PYTHON_COMMAND" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else "Python 3.10 or newer is required")'

if [ ! -x "$PROJECT_ROOT/.venv/bin/python" ]; then
    echo "Creating Python virtual environment: $PROJECT_ROOT/.venv"
    "$PYTHON_COMMAND" -m venv "$PROJECT_ROOT/.venv"
fi

VENV_PYTHON="$PROJECT_ROOT/.venv/bin/python"
echo "Installing/updating dependencies..."
"$VENV_PYTHON" -m pip install --disable-pip-version-check --upgrade pip
"$VENV_PYTHON" -m pip install --disable-pip-version-check -r "$PROJECT_ROOT/requirements.txt"
"$VENV_PYTHON" -c "import docx, fitz, pypandoc, pypdf, requests, urllib3; print('Dependency check passed.')"
echo "Installation completed."
