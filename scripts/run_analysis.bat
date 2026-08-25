@echo off
rem ============================================================
rem  ECOCO complaint analysis - scheduled job
rem  Called by Windows Task Scheduler (see register_task.ps1).
rem  NOTE: keep this file ASCII-only. Windows parses .bat with the
rem  OEM codepage, so UTF-8 Chinese comments break the parser.
rem
rem  Source mode is picked automatically:
rem    SOURCE_SHEET_ID set    -> read that Google Sheet (--from-sheet)
rem    otherwise              -> watch the inbox folder (ECOCO_INBOX)
rem  This avoids the silent-no-op failure mode where --from-sheet is
rem  scheduled but no source sheet is configured.
rem
rem  Exit codes: 0 = done, 2 = no new data (treated as success),
rem              anything else = failure
rem ============================================================
setlocal

rem UTF-8 console + Python output, otherwise Chinese in the log is garbled
chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

set "PROJECT_DIR=%~dp0.."
pushd "%PROJECT_DIR%"

rem Load local environment variables if present
if exist "scripts\env.local.bat" call "scripts\env.local.bat"

rem Python interpreter: override with ECOCO_PYTHON if needed
if "%ECOCO_PYTHON%"=="" set "ECOCO_PYTHON=python"

rem Folder watched when no source sheet is configured
if "%ECOCO_INBOX%"=="" set "ECOCO_INBOX=inbox"

rem Locale-independent timestamps. %DATE% is localized (on zh-TW it starts
rem with the weekday), which would corrupt the log file name.
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "TODAY=%%i"
for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm:ss\""') do set "NOW=%%i"

if not exist "logs" mkdir "logs"
set "LOGFILE=logs\analysis_%TODAY%.log"

echo.>> "%LOGFILE%"
echo ============================================================>> "%LOGFILE%"
echo [%NOW%] start>> "%LOGFILE%"

if not "%SOURCE_SHEET_ID%"=="" (
    echo [%NOW%] source: google sheet %SOURCE_SHEET_ID%>> "%LOGFILE%"
    "%ECOCO_PYTHON%" -m automation.cli run --from-sheet --only-new --report --out output >> "%LOGFILE%" 2>&1
) else (
    if not exist "%ECOCO_INBOX%" mkdir "%ECOCO_INBOX%"
    echo [%NOW%] source: folder %ECOCO_INBOX%>> "%LOGFILE%"
    "%ECOCO_PYTHON%" -m automation.cli watch --dir "%ECOCO_INBOX%" --out output --report >> "%LOGFILE%" 2>&1
)
set "RC=%ERRORLEVEL%"

for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm:ss\""') do set "NOW=%%i"
if "%RC%"=="0" (
    echo [%NOW%] done>> "%LOGFILE%"
) else if "%RC%"=="2" (
    echo [%NOW%] no new data, skipped>> "%LOGFILE%"
    set "RC=0"
) else (
    echo [%NOW%] FAILED with exit code %RC%>> "%LOGFILE%"
)

popd
endlocal & exit /b %RC%
