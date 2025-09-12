@echo off
echo ========================================
echo  PM Form Generator - Build Script
echo ========================================
echo.

echo [1/4] Installing requirements...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install requirements
    pause
    exit /b 1
)

echo.
echo [2/4] Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"

echo.
echo [3/4] Building executable...
pyinstaller --onefile --windowed --icon=formgenerator.ico --name="PM_Form_Generator" formgenerator.py
if %errorlevel% neq 0 (
    echo ERROR: Failed to build executable
    pause
    exit /b 1
)

echo.
echo [4/4] Copying files to dist folder...
if not exist "dist\assets" mkdir "dist\assets"
copy "README.md" "dist\" >nul 2>&1
copy "ui.html" "dist\assets\" >nul 2>&1

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Executable location: dist\PM_Form_Generator.exe
echo Documentation: dist\README.md
echo Visual Guide: dist\assets\ui.html
echo.
echo You can now distribute the 'dist' folder
echo or just the PM_Form_Generator.exe file.
echo.
pause