#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

echo "========================================"
echo "Registration to Registry Launcher"
echo "========================================"
echo

if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_CMD="python"
else
    echo "Python was not found on this computer."
    echo "Please install Python first, then run this file again."
    exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
    echo "Creating virtual environment..."
    "$PYTHON_CMD" -m venv .venv
fi

echo "Installing required library..."
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt

echo
 echo "Running script..."
".venv/bin/python" registration_to_registry.py

echo
 echo "Script finished."
