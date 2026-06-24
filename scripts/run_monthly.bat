@echo off
REM ============================================================
REM  StationDeck -- Monthly Report Generator
REM  Double-click OR run via Task Scheduler.
REM ============================================================

REM -- Step 1: Move to the project root (one level above scripts\)
cd /d "%~dp0.."

REM -- Step 2: Detect current month and year using PowerShell
for /f %%A in ('powershell -NoProfile -Command "Get-Date -Format MM"') do set MONTH=%%A
for /f %%A in ('powershell -NoProfile -Command "Get-Date -Format yyyy"') do set YEAR=%%A

REM -- Remove leading zero from month (so "05" becomes "5")
set /a MONTH_NUM=%MONTH%

REM -- Step 3: Set log file path
set LOGFILE=%~dp0..\logs\scheduler.log

REM -- Step 4: Activate the virtual environment
call venv\Scripts\activate.bat

REM -- Step 5: Write header to scheduler.log
echo. >> "%LOGFILE%"
echo ======================================================== >> "%LOGFILE%"
echo  StationDeck -- Scheduler Run >> "%LOGFILE%"
echo  Month: %MONTH_NUM%   Year: %YEAR% >> "%LOGFILE%"
for /f "delims=" %%T in ('powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-dd HH:mm:ss'"') do echo  Started: %%T >> "%LOGFILE%"
echo ======================================================== >> "%LOGFILE%"

REM -- Step 6: Run the report generator, redirect all output to scheduler.log
python main.py --month %MONTH_NUM% --year %YEAR% >> "%LOGFILE%" 2>&1

REM -- Step 7: Write result to scheduler.log
if %ERRORLEVEL% EQU 0 (
    echo  RESULT: COMPLETED successfully. >> "%LOGFILE%"
) else (
    echo  RESULT: FAILED. See logs\app.log for details. >> "%LOGFILE%"
)
echo ======================================================== >> "%LOGFILE%"

REM -- Step 8: Also show output in terminal if run manually
echo.
echo  Done. Output written to logs\scheduler.log
echo.
cmd /k