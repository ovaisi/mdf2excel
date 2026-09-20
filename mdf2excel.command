#!/bin/bash
# MDF to Excel Converter - double-click launcher for macOS.
cd "$(dirname "$0")"
bash mdf2excel.sh
STATUS=$?
if [ $STATUS -ne 0 ]; then
    echo ""
    echo "Exited with an error (code $STATUS). Press Enter to close this window."
    read
fi
