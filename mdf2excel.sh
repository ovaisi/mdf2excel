#!/usr/bin/env bash
# MDF to Excel Converter - macOS / Linux launcher.
#
# Usage:
#   ./mdf2excel.sh                    # open the graphical window
#   ./mdf2excel.sh "database.mdf"     # convert a file (command line)
#   ./mdf2excel.sh "database.mdf" --tables
#   ./mdf2excel.sh "database.mdf" --info
#
# Make it executable if needed:  chmod +x mdf2excel.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

PY=""
for c in python3 python; do
    if command -v "$c" >/dev/null 2>&1; then
        PY="$c"
        break
    fi
done

if [ -z "$PY" ]; then
    echo "Error: Python 3 was not found. Install Python 3 first." >&2
    exit 1
fi

if ! "$PY" -c "import openpyxl" >/dev/null 2>&1; then
    echo "Installing the openpyxl dependency..."
    "$PY" -m pip install --user openpyxl 2>/dev/null || "$PY" -m pip install openpyxl
fi

exec "$PY" mdf2excel.py "$@"
