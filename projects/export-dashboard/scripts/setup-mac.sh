#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
PYTHON_CMD="${PYTHON_CMD:-python3}"
SKIP_WEB_DEPENDENCIES=0

while [[ $# -gt 0 ]]; do
  case "$1" in
    --skip-web-dependencies)
      SKIP_WEB_DEPENDENCIES=1
      shift
      ;;
    --python)
      PYTHON_CMD="${2:?Missing value for --python}"
      shift 2
      ;;
    *)
      echo "Unknown argument: $1" >&2
      exit 2
      ;;
  esac
done

"$PYTHON_CMD" -m venv "$REPO_ROOT/.venv"
PYTHON="$REPO_ROOT/.venv/bin/python"
"$PYTHON" -m pip install --upgrade pip
"$PYTHON" -m pip install -r "$REPO_ROOT/requirements.txt"

if [[ "$SKIP_WEB_DEPENDENCIES" -eq 0 ]]; then
  if ! command -v npm >/dev/null 2>&1; then
    echo "Node.js 20+ and npm are required for the Vercel application." >&2
    exit 1
  fi
  (cd "$REPO_ROOT/web" && npm install)
fi

echo "Environment ready: $REPO_ROOT/.venv"
