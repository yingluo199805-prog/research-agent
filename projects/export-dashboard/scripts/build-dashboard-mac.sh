#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
BUILDER_ROOT="$REPO_ROOT/src/dashboard_build"

DATA_WORKBOOK=""
DATA_FILE_DATE="${DASHBOARD_DATA_FILE_DATE:-}"
MKLS_MAX_YEAR="${DASHBOARD_MKLS_MAX_YEAR:-}"
MKLS_MAX_MONTH="${DASHBOARD_MKLS_MAX_MONTH:-}"

usage() {
  cat <<'EOF'
Usage:
  scripts/build-dashboard-mac.sh --data-workbook <master.xlsx> [options]

Options:
  --data-file-date YYYY-MM-DD   Override data update date. If omitted, a YYMMDD
                                stamp is read from the workbook filename.
  --mkls-max-year YYYY          Calendar year to cap for MKLS monthly data.
  --mkls-max-month N            Keep only months 1..N for the capped MKLS year.

Example:
  DATA_WORKBOOK="$(find "$HOME/Library/CloudStorage" -path '*/work/data/export-dashboard-data/master/master_260702_working.xlsx' -print -quit)"
  scripts/build-dashboard-mac.sh \
    --data-workbook "$DATA_WORKBOOK" \
    --data-file-date 2026-07-02 \
    --mkls-max-year 2026 \
    --mkls-max-month 5
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --data-workbook)
      DATA_WORKBOOK="${2:?Missing value for --data-workbook}"
      shift 2
      ;;
    --data-file-date)
      DATA_FILE_DATE="${2:?Missing value for --data-file-date}"
      shift 2
      ;;
    --mkls-max-year)
      MKLS_MAX_YEAR="${2:?Missing value for --mkls-max-year}"
      shift 2
      ;;
    --mkls-max-month)
      MKLS_MAX_MONTH="${2:?Missing value for --mkls-max-month}"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

if [[ -z "$DATA_WORKBOOK" ]]; then
  echo "--data-workbook is required." >&2
  usage >&2
  exit 2
fi

PYTHON="$REPO_ROOT/.venv/bin/python"
if [[ ! -x "$PYTHON" ]]; then
  PYTHON="${PYTHON_CMD:-python3}"
fi

DATA_WORKBOOK_ABS="$("$PYTHON" -c 'from pathlib import Path; import sys; print(Path(sys.argv[1]).expanduser().resolve())' "$DATA_WORKBOOK")"
if [[ ! -f "$DATA_WORKBOOK_ABS" ]]; then
  echo "Data workbook not found: $DATA_WORKBOOK_ABS" >&2
  exit 1
fi

if [[ -z "$DATA_FILE_DATE" ]]; then
  filename="$(basename "$DATA_WORKBOOK_ABS")"
  if [[ "$filename" =~ ([0-9]{6}) ]]; then
    stamp="${BASH_REMATCH[1]}"
    DATA_FILE_DATE="20${stamp:0:2}-${stamp:2:2}-${stamp:4:2}"
  else
    echo "Use --data-file-date YYYY-MM-DD when the workbook name has no YYMMDD stamp." >&2
    exit 1
  fi
fi

export DASHBOARD_DATA_FILE="$DATA_WORKBOOK_ABS"
export DASHBOARD_DATA_FILE_DATE="$DATA_FILE_DATE"
if [[ -n "$MKLS_MAX_YEAR" ]]; then export DASHBOARD_MKLS_MAX_YEAR="$MKLS_MAX_YEAR"; fi
if [[ -n "$MKLS_MAX_MONTH" ]]; then export DASHBOARD_MKLS_MAX_MONTH="$MKLS_MAX_MONTH"; fi

(cd "$BUILDER_ROOT" && \
  "$PYTHON" gen_data.py && \
  "$PYTHON" gen_data_caam.py && \
  "$PYTHON" gen_data_mkls.py && \
  "$PYTHON" build.py)

HTML="$BUILDER_ROOT/index.html"
BYTES="$("$PYTHON" -c 'from pathlib import Path; import sys; print(Path(sys.argv[1]).stat().st_size)' "$HTML")"
if [[ "$BYTES" -lt 1048576 ]]; then
  echo "Dashboard HTML was not generated or is unexpectedly small: $BYTES bytes" >&2
  exit 1
fi

mkdir -p "$REPO_ROOT/web/public"
cp "$HTML" "$REPO_ROOT/web/public/dashboard.html"

echo "Built dashboard:"
echo "  source: $HTML ($BYTES bytes)"
echo "  target: $REPO_ROOT/web/public/dashboard.html"
