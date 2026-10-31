#!/usr/bin/env bash
#
# run.sh — CRediT contributor-role parser (QE Hackathon)
#
# Usage:
#   ./run.sh <input_folder> <output.csv>
#
# Runs cleanly on a bare Ubuntu environment with Python 3.12 and internet
# access. Installs all required Python dependencies, then runs the pipeline.
#
set -euo pipefail

if [ "$#" -ne 2 ]; then
    echo "Usage: ./run.sh <input_folder> <output.csv>" >&2
    exit 1
fi

INPUT_FOLDER="$1"
OUTPUT_CSV="$2"

# Resolve the directory this script lives in, so it works from any CWD.
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Self-install dependencies (quiet). Prefer python3; fall back to python.
PY="$(command -v python3 || command -v python)"
echo "[run.sh] Using interpreter: $PY"
echo "[run.sh] Installing dependencies..."
"$PY" -m pip install --quiet --upgrade \
    requests beautifulsoup4 python-docx rapidfuzz

echo "[run.sh] Running parser: $INPUT_FOLDER -> $OUTPUT_CSV"
"$PY" "$SCRIPT_DIR/main.py" "$INPUT_FOLDER" "$OUTPUT_CSV"

echo "[run.sh] Done."
