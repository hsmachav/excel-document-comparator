@echo off
echo.
echo ========================================
echo Excel Document Comparator - Build Script
echo ========================================
echo.

echo Step 1: Checking Python installation...
python --version
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.7+
    pause
    exit /b 1
)

echo.
echo Step 2: Installing/Updating PyInstaller...
pip install --upgrade pyinstaller
if errorlevel 1 (
    echo ERROR: Failed to install PyInstaller
    pause
    exit /b 1
)

echo.
echo Step 3: Building executable...
pyinstaller --onefile --windowed --name=ExcelDocumentComparator --distpath=dist --buildpath=build --specpath=. src\main.py

if errorlevel 1 (
    echo ERROR: Failed to build executable
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Executable location:
echo dist\ExcelDocumentComparator.exe
echo.
pause
