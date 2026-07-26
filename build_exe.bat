@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building executable...
pyinstaller --onefile --windowed --icon=icon.ico --name=ExcelDocumentComparator src\main.py

echo.
echo Build complete! Executable is in the dist\ folder.
pause