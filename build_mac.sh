#!/bin/bash
echo "[1/2] Installing dependencies..."
pip3 install openpyxl pandas pyinstaller

STAMP=$(date +%y%m%d_%H%M)
FNAME="MUI_tool_1_4_${STAMP}"

echo "[2/2] Building executable: $FNAME"
sed "s/BUILD_DATETIME/${STAMP}/" translation_gui.py > translation_gui_build.py
pyinstaller --onefile --windowed --name "$FNAME" --hidden-import openpyxl --hidden-import pandas translation_gui_build.py
rm -f translation_gui_build.py

echo "Done! Find $FNAME in the dist folder."
