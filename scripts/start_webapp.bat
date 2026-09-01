@echo off
rem ============================================================
rem  ECOCO complaint analysis - start the web UI
rem  Double-click, or run from a terminal:
rem      scripts\start_webapp.bat
rem  Keep this window open; closing it stops the server.
rem  NOTE: keep this file ASCII-only. Windows parses .bat with the
rem  OEM codepage, so UTF-8 Chinese comments break the parser.
rem ============================================================
setlocal

chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

set "PROJECT_DIR=%~dp0.."
pushd "%PROJECT_DIR%"

if exist "scripts\env.local.bat" call "scripts\env.local.bat"

if "%ECOCO_PYTHON%"=="" set "ECOCO_PYTHON=python"
if "%ECOCO_PORT%"==""   set "ECOCO_PORT=8501"

rem Bind 0.0.0.0 so other machines on the LAN can reach it too
echo.
echo   ECOCO complaint analysis - web UI
echo   ---------------------------------------------------------
echo   Local   : http://localhost:%ECOCO_PORT%
echo   Network : http://%COMPUTERNAME%:%ECOCO_PORT%
echo.
echo   Press Ctrl+C to stop.
echo.

"%ECOCO_PYTHON%" -m streamlit run complaint_webapp.py ^
    --server.port %ECOCO_PORT% ^
    --server.address 0.0.0.0 ^
    --server.headless true ^
    --browser.gatherUsageStats false

popd
endlocal
