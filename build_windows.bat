@echo off
echo RC Generator - Windows Build Script
echo =====================================

echo Installing Python dependencies...
pip install pandas numpy PyQt6 xlsxwriter pyinstaller

echo Building Windows executable...
pyinstaller --onefile --windowed --name "RC_Generator" ^
    --add-data "Planificare Elmet.xlsx;." ^
    --add-data "Tehnologii.xlsx;." ^
    --version-file "version_info.txt" ^
    --manifest "manifest.xml" ^
    route_card_coc_app.py

echo.
echo Build complete! 
echo Executable location: dist\RC_Generator.exe
echo.
pause
