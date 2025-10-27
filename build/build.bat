@echo off
REM Build script for SharePoint Permissions Exceler GUI executable
REM This script creates a standalone .exe file that includes all dependencies

echo ========================================
echo SharePoint Permissions Exceler Builder
echo ========================================
echo.

REM Check if we're in the build directory
if not exist "gui.spec" (
    echo ERROR: gui.spec not found. Make sure you're running this from the build directory.
    pause
    exit /b 1
)

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.7+ and try again.
    pause
    exit /b 1
)

REM Check if we're in a virtual environment (recommended)
python -c "import sys; exit(0 if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix) else 1)" >nul 2>&1
if errorlevel 1 (
    echo WARNING: Not running in a virtual environment.
    echo It's recommended to activate your virtual environment first.
    echo.
    set /p continue="Continue anyway? (y/N): "
    if /i not "%continue%"=="y" (
        echo Build cancelled.
        pause
        exit /b 1
    )
    echo.
)

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller.
        pause
        exit /b 1
    )
)

REM Check if all required dependencies are installed
echo Checking dependencies...
python -c "import pandas, openpyxl, msal, requests, dotenv; from PyQt6.QtWidgets import QApplication" >nul 2>&1
if errorlevel 1 (
    echo ERROR: Missing required dependencies.
    echo Please run: pip install -r ../requirements.txt
    pause
    exit /b 1
)

echo Dependencies check passed.
echo.

REM Clean previous build artifacts
echo Cleaning previous build artifacts...
if exist "dist" rmdir /s /q "dist"
if exist "build_temp" rmdir /s /q "build_temp"

REM Create output directory
if not exist "dist" mkdir "dist"

echo.
echo Building executable...
echo This may take several minutes...
echo.

REM Run PyInstaller with our spec file
pyinstaller --distpath="dist" --workpath="build_temp" gui.spec

if errorlevel 1 (
    echo.
    echo ERROR: Build failed!
    echo Check the output above for error details.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Executable location: build\dist\SharePoint-Permissions-Exceler.exe
echo.
echo You can now distribute this .exe file to users who don't have Python installed.
echo The executable includes all required dependencies.
echo.

REM Check if the executable was actually created
if exist "dist\SharePoint-Permissions-Exceler.exe" (
    echo File size:
    dir "dist\SharePoint-Permissions-Exceler.exe" | findstr "SharePoint-Permissions-Exceler.exe"
    echo.
    echo The executable is ready for distribution!
) else (
    echo WARNING: Executable file not found in expected location.
    echo Check the dist directory for the output file.
)

echo.
pause