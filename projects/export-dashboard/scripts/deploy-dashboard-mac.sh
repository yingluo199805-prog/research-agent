#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

SCOPE="${VERCEL_SCOPE:-}"
BUILD_ARGS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    --scope)
      SCOPE="${2:?Missing value for --scope}"
      shift 2
      ;;
    --data-workbook|--data-file-date|--mkls-max-year|--mkls-max-month)
      BUILD_ARGS+=("$1" "${2:?Missing value for $1}")
      shift 2
      ;;
    -h|--help)
      "$SCRIPT_DIR/build-dashboard-mac.sh" --help
      echo
      echo "Additional deploy option:"
      echo "  --scope <vercel-scope>"
      exit 0
      ;;
    *)
      echo "Unknown argument: $1" >&2
      exit 2
      ;;
  esac
done

"$SCRIPT_DIR/build-dashboard-mac.sh" "${BUILD_ARGS[@]}"

if ! command -v vercel >/dev/null 2>&1; then
  echo "Install and authenticate the Vercel CLI before deployment." >&2
  exit 1
fi

export NODE_USE_ENV_PROXY="${NODE_USE_ENV_PROXY:-1}"
VERCEL_ARGS=(--prod --yes)
if [[ -n "$SCOPE" ]]; then
  VERCEL_ARGS+=(--scope "$SCOPE")
fi

(cd "$REPO_ROOT/web" && vercel "${VERCEL_ARGS[@]}")
