@echo off
rem ============================================================
rem  ECOCO complaint analysis - scheduled job
rem  Called by Windows Task Scheduler (see register_task.ps1).
rem  NOTE: keep this file ASCII-only. Windows parses .bat with the
rem  OEM codepage, so UTF-8 Chinese comments break the parser.
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

rem Locale-independent timestamps. %DATE% is localized (on zh-TW it starts
rem with the weekday), which would corrupt the log file name.
for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set "TODAY=%%i"
for /f "delims=" %%i in ('powershell -NoProfile -Command "Get-Date -Format \"yyyy-MM-dd HH:mm:ss\""') do set "NOW=%%i"

if not exist "logs" mkdir "logs"
set "LOGFILE=logs\analysis_%TODAY%.log"

echo.>> "%LOGFILE%"
echo ============================================================>> "%LOGFILE%"
echo [%NOW%] start>> "%LOGFILE%"

rem Main job: read new rows from the source sheet, classify, write report.
rem For a shared-folder source instead, use:
rem   "%ECOCO_PYTHON%" -m automation.cli watch --dir "\server\share\inbox" --out output --report
"%ECOCO_PYTHON%" -m automation.cli run --from-sheet --only-new --report --out output >> "%LOGFILE%" 2>&1
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
