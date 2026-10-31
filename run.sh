#!/usr/bin/env bash
#
# run.sh — CRediT contributor-role parser (QE Hackathon, Problem 2)
#
# Usage:
#   ./run.sh <input_folder> <output.csv>
#
# Designed to work on a clean Ubuntu machine with Python 3.12 and internet
# access. It installs all required Python dependencies — bootstrapping pip if
# the machine lacks it, and handling Ubuntu's PEP 668 "externally-managed-
# environment" restriction — then runs the pipeline.
#
set -euo pipefail

if [ "$#" -ne 2 ]; then
    echo "Usage: ./run.sh <input_folder> <output.csv>" >&2
    exit 1
fi

INPUT_FOLDER="$1"
OUTPUT_CSV="$2"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Pick an interpreter (prefer python3).
PY="$(command -v python3 || command -v python)"
echo "[run.sh] Using interpreter: $PY"
"$PY" --version

# Use sudo for apt only if we are not already root and sudo exists.
SUDO=""
if [ "$(id -u)" -ne 0 ] && command -v sudo >/dev/null 2>&1; then
    SUDO="sudo"
fi

# 1) Ensure pip is available (clean Ubuntu may ship Python without pip).
if ! "$PY" -m pip --version >/dev/null 2>&1; then
    echo "[run.sh] pip not found; bootstrapping with ensurepip..."
    "$PY" -m ensurepip --upgrade >/dev/null 2>&1 || true
fi
if ! "$PY" -m pip --version >/dev/null 2>&1; then
    if command -v apt-get >/dev/null 2>&1; then
        echo "[run.sh] Installing python3-pip via apt-get..."
        $SUDO apt-get update -y
        $SUDO apt-get install -y python3-pip
    fi
fi

# 2) Install dependencies. Try a normal install first; fall back to
#    --break-system-packages (PEP 668 on Ubuntu 23.04+), then to --user.
DEPS="requests beautifulsoup4 python-docx rapidfuzz"
echo "[run.sh] Installing dependencies: $DEPS"
"$PY" -m pip install --quiet --upgrade $DEPS \
    || "$PY" -m pip install --quiet --upgrade --break-system-packages $DEPS \
    || "$PY" -m pip install --quiet --upgrade --user $DEPS

# 3) Run the pipeline.
echo "[run.sh] Running parser: $INPUT_FOLDER -> $OUTPUT_CSV"
"$PY" "$SCRIPT_DIR/main.py" "$INPUT_FOLDER" "$OUTPUT_CSV"

echo "[run.sh] Done. Wrote $OUTPUT_CSV"
