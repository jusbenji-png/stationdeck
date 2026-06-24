@echo off
REM =============================================================
REM  StationDeck — Windows Build Script
REM  Double-click this file to build the installer
REM
REM  Requirements:
REM    - Run from the project root: C:\Users\LENOVO\stationdeck\
REM    - Virtual environment must be activated OR pyinstaller
REM      must be installed in the venv
REM
REM  Output: dist\StationDeck\StationDeck.exe
REM =============================================================

echo.
echo ============================================================
echo   StationDeck Build Script
echo   Building Windows .exe package...
echo ============================================================
echo.

REM ── Activate virtual environment ────────────────────────────
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Could not activate virtual environment.
    echo Make sure you are running this from C:\Users\LENOVO\stationdeck\
    pause
    exit /b 1
)

REM ── Install PyInstaller if not already installed ─────────────
echo [1/4] Checking PyInstaller...
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo     Installing PyInstaller...
    pip install pyinstaller
)
echo     PyInstaller ready.
echo.

REM ── Clean previous build artifacts ──────────────────────────
echo [2/4] Cleaning previous build...
if exist build\StationDeck rmdir /s /q build\StationDeck
if exist dist\StationDeck  rmdir /s /q dist\StationDeck
echo     Clean complete.
echo.

REM ── Run PyInstaller ──────────────────────────────────────────
echo [3/4] Running PyInstaller...
echo     This takes 2-5 minutes. Please wait...
echo.
pyinstaller stationdeck.spec --noconfirm
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed.
    echo Check the output above for error details.
    pause
    exit /b 1
)
echo.

REM ── Copy external files to dist folder ───────────────────────
REM  These files must live next to the .exe, NOT inside the bundle.
REM  They are station-specific and should not be overwritten on update.
echo [4/4] Copying external files to dist folder...

REM Create required folders in dist
mkdir dist\StationDeck\data\input      2>nul
mkdir dist\StationDeck\data\processed  2>nul
mkdir dist\StationDeck\data\ocr_temp   2>nul
mkdir dist\StationDeck\data\ocr_audit  2>nul
mkdir dist\StationDeck\config\stations 2>nul
mkdir dist\StationDeck\reports         2>nul
mkdir dist\StationDeck\logs            2>nul
mkdir dist\StationDeck\installer       2>nul

REM Copy station config (YAML)
copy "config\stations\te_rwizi.yaml" "dist\StationDeck\config\stations\te_rwizi.yaml" >nul
echo     Copied: config\stations\te_rwizi.yaml

REM Write version file directly from AppVersion — never gets out of sync
set STATIONDECK_VERSION=1.0.2
echo %STATIONDECK_VERSION%> "dist\StationDeck\config\version.txt"
echo     Written: config\version.txt (%STATIONDECK_VERSION%)

REM Copy .env template if .env doesn't already exist in dist
if not exist "dist\StationDeck\.env" (
    (
        echo # StationDeck Environment Configuration
        echo # Fill in your values below
        echo.
        echo OPENAI_API_KEY=your_openai_key_here
        echo STATION_NAME=Your Station Name
        echo STATION_LOCATION=Your Location, Uganda
        echo SMTP_HOST=smtp.gmail.com
        echo SMTP_PORT=587
        echo SMTP_USER=your_email@gmail.com
        echo SMTP_PASS=your_app_password
        echo REPORT_RECIPIENTS=manager@station.com
        echo STATION_PASSWORD=stationdeck123
    ) > "dist\StationDeck\.env"
    echo     Created: .env template
)

echo.
echo ============================================================
echo   BUILD COMPLETE
echo.
echo   Output folder:  dist\StationDeck\
echo   Executable:     dist\StationDeck\StationDeck.exe
echo.
echo   NEXT STEP:
echo   Open StationDeck.iss in Inno Setup and press F9
echo   to build the installer: installer\StationDeck_Setup.exe
echo ============================================================
echo.
pause