#!/usr/bin/env bash
set -euo pipefail

# Builds FinanceSummary and places it in THIS SAME FOLDER.
# Run this file on macOS.

python3 -m venv .venv
source .venv/bin/activate

python -m pip install --upgrade pip
pip install -r requirements.txt

pyinstaller --clean --noconfirm --onefile --name FinanceSummary --distpath . build_accounting_summary.py

echo
echo "Build complete."
echo "Your executable is now in this folder:"
echo "FinanceSummary"
echo
echo "You can move FinanceSummary into the folder with your Daily Finances Excel files."
