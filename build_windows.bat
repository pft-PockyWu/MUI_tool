@echo off
chcp 65001 >nul
echo [1/2] Installing dependencies...
pip install openpyxl pandas pyinstaller

powershell -NoProfile -Command "Get-Date -Format 'yyMMdd_HHmm'" > %TEMP%\mui_stamp.tmp
set /p STAMP=< %TEMP%\mui_stamp.tmp
del %TEMP%\mui_stamp.tmp 2>nul
set FNAME=MUI_tool_2_4_%STAMP%

echo [2/2] Building executable...
powershell -Command "(Get-Content translation_gui.py -Encoding UTF8) -replace 'BUILD_DATETIME', '%STAMP%' | Set-Content translation_gui_build.py -Encoding UTF8"
python -m PyInstaller --onefile --windowed --name "%FNAME%" --hidden-import openpyxl --hidden-import pandas translation_gui_build.py
del translation_gui_build.py 2>nul

echo Done! Find %FNAME%.exe in the dist folder.
echo.
set /p UPLOAD="Upload to Google Drive? (Y/N): "
if /i "%UPLOAD%"=="Y" (
    echo [3/3] Uploading to Google Drive...
    python upload_to_drive.py
)
pause
