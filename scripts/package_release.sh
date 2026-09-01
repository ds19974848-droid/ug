#!/usr/bin/env bash
set -euo pipefail

# scripts/package_release.sh
# Create a release tarball for the current branch after running tests and optional DB backfill dry-run.
# Usage:
#   ./scripts/package_release.sh --branch fix/match-improvements --output ../ug-release.tar.gz
#   ./scripts/package_release.sh --branch fix/match-improvements --output ../ug-release.tar.gz --db sqlite:///path/to/db

usage() {
  echo "Usage: $0 --branch BRANCH --output OUTPUT_TAR_GZ [--db DATABASE_URL] [--venv venv_dir]"
  exit 2
}

BRANCH=""
OUT=""
DB_URL=""
VENV_DIR=".venv"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --branch) BRANCH="$2"; shift 2;;
    --output) OUT="$2"; shift 2;;
    --db) DB_URL="$2"; shift 2;;
    --venv) VENV_DIR="$2"; shift 2;;
    -h|--help) usage;;
    *) echo "Unknown arg: $1"; usage;;
  esac
done

if [[ -z "$BRANCH" || -z "$OUT" ]]; then
  usage
fi

echo "Packaging branch: $BRANCH -> $OUT"

# Ensure we are in repo root (script expects to run from repo root)
REPO_ROOT=$(git rev-parse --show-toplevel 2>/dev/null || true)
if [[ -z "$REPO_ROOT" ]]; then
  echo "Not in a git repository. Please run this script from the repository root." >&2
  exit 1
fi
cd "$REPO_ROOT"

# Fetch and checkout branch
git fetch origin "$BRANCH":"$BRANCH" || true
git checkout "$BRANCH"

# Create venv and install dependencies
python -m pip install --upgrade pip
python -m pip install virtualenv >/dev/null 2>&1 || true
if [[ ! -d "$VENV_DIR" ]]; then
  python -m venv "$VENV_DIR"
fi
# shellcheck source=/dev/null
source "$VENV_DIR/bin/activate"

if [[ -f requirements.txt ]]; then
  pip install -r requirements.txt || true
fi
pip install pytest rapidfuzz || true

# Run unit tests (only the matching tests by default)
echo "Running unit tests..."
pytest -q || {
  echo "Tests failed. Aborting packaging." >&2
  deactivate || true
  exit 1
}

echo "Tests passed. Proceeding to optional DB backfill dry-run (if DB provided)."
if [[ -n "$DB_URL" ]]; then
  echo "Running backfill dry-run against: $DB_URL"
  # Run backfill in dry-run mode if supported
  python scripts/backfill_normalized.py --dry-run --db "$DB_URL" || {
    echo "Backfill dry-run failed. Aborting." >&2
    deactivate || true
    exit 1
  }
fi

# Create release tarball using git archive so untracked files are not included
TMP_OUT="$(mktemp -u)"

git archive --format=tar.gz --output="$OUT" "$BRANCH"

SIZE=$(stat -c%s "$OUT" 2>/dev/null || stat -f%z "$OUT")
if [[ -z "$SIZE" ]]; then SIZE=0; fi

echo "Created package: $OUT (size: $SIZE bytes)"

deactivate || true

echo "Done. If you want a Docker image, build with: docker build -t ug:release -f Dockerfile ."
