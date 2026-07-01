#!/bin/bash
set -e
echo "Building sap-cases (macOS) ..."

if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found. Install from https://www.python.org/downloads/"
    exit 1
fi

echo "Installing dependencies..."
pip3 install pyinstaller playwright requests openpyxl pyautogui
playwright install chromium

echo "Running PyInstaller..."
pyinstaller --onefile --name sap-cases --clean sap_cases.py

echo ""
echo "Done! Executable: dist/sap-cases"
echo ""
echo "Distribute to colleagues:"
echo "  dist/sap-cases  +  config.json  (same folder)"
echo ""
echo "First run on each machine:"
echo "  ./sap-cases --customer 450162 --relogin"
echo "  ./sap-cases setup-cron"
