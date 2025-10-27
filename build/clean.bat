@echo off
REM Clean script for SharePoint Permissions Exceler build artifacts
REM Removes all build-related temporary files and directories

echo ========================================
echo SharePoint Permissions Exceler Cleaner
echo ========================================
echo.

REM Check if we're in the build directory
if not exist "gui.spec" (
    echo ERROR: gui.spec not found. Make sure you're running this from the build directory.
    pause
    exit /b 1
)

echo Cleaning build artifacts...
echo.

REM Remove PyInstaller build directories
if exist "dist" (
    echo Removing dist directory...
    rmdir /s /q "dist"
)

if exist "build_temp" (
    echo Removing build_temp directory...
    rmdir /s /q "build_temp"
)

REM Remove PyInstaller cache (if any)
if exist "__pycache__" (
    echo Removing __pycache__ directory...
    rmdir /s /q "__pycache__"
)

REM Remove any .pyc files
if exist "*.pyc" (
    echo Removing .pyc files...
    del /q "*.pyc"
)

REM Clean parent directory artifacts as well
cd ..

REM Remove Python cache directories in project root
if exist "__pycache__" (
    echo Removing project __pycache__ directory...
    rmdir /s /q "__pycache__"
)

REM Remove any .pyc files in project root
if exist "*.pyc" (
    echo Removing project .pyc files...
    del /q "*.pyc"
)

cd build

echo.
echo ========================================
echo Cleanup completed!
echo ========================================
echo.
echo All build artifacts have been removed.
echo You can now run build.bat to create a fresh build.
echo.
pause