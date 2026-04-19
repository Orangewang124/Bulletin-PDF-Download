@echo off

cd /d "%~dp0"

echo ========================================
echo   PDF Bulletin Downloader - Build EXE
echo ========================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python and add to PATH.
    pause
    exit /b 1
)

python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] PyInstaller not installed, installing...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo [ERROR] PyInstaller install failed. Run manually: pip install pyinstaller
        pause
        exit /b 1
    )
)

if not exist "MoonOrangeBulletinPDFDownloader.spec" (
    echo [ERROR] MoonOrangeBulletinPDFDownloader.spec not found.
    pause
    exit /b 1
)

echo [INFO] Building...
echo.

pyinstaller MoonOrangeBulletinPDFDownloader.spec --clean

if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo   Build SUCCESS!
    echo   Output: dist\MoonOrangeBulletinPDFDownloader.exe
    echo ========================================
) else (
    echo.
    echo [ERROR] Build failed. Check error messages above.
)

pause
