@echo off
echo ================================================================================
echo Excel Cleaner - Build Script
echo ================================================================================
echo.

echo Step 1: Installing dependencies...
echo.
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install dependencies
    echo Please ensure Python and pip are installed correctly
    pause
    exit /b 1
)

echo.
echo ================================================================================
echo Step 2: Building executable...
echo ================================================================================
echo.
pyinstaller excel_cleaner.spec
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed
    echo Trying alternative method...
    python -m PyInstaller excel_cleaner.spec
    if %errorlevel% neq 0 (
        echo.
        echo ERROR: Build failed with both methods
        pause
        exit /b 1
    )
)

echo.
echo ================================================================================
echo BUILD SUCCESSFUL!
echo ================================================================================
echo.
echo Your executable is ready at: dist\ExcelCleaner.exe
echo.
echo You can now:
echo   1. Copy dist\ExcelCleaner.exe to any location
echo   2. Drag and drop Excel files onto it
echo   3. Or double-click it to browse for files
echo.
echo ================================================================================
pause
