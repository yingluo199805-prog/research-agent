#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
PYTHON="$REPO_ROOT/.venv/bin/python"
if [[ ! -x "$PYTHON" ]]; then
  PYTHON="${PYTHON_CMD:-python3}"
fi

(cd "$REPO_ROOT" && "$PYTHON" -X utf8 - <<'PY'
import ast
from pathlib import Path

files = list(Path("src").rglob("*.py")) + list(Path("archive").rglob("*.py"))
for path in files:
    ast.parse(path.read_text(encoding="utf-8-sig"))
print(f"Python syntax: OK ({len(files)} files)")
PY
)

if command -v node >/dev/null 2>&1; then
  (cd "$REPO_ROOT/web" && node -e "JSON.parse(require('fs').readFileSync('package.json', 'utf8')); console.log('Node package metadata: OK')")
else
  echo "node not found; skipped web package metadata check."
fi
