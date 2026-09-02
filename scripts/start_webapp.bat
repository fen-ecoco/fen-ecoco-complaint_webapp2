@echo off
rem ============================================================
rem  ECOCO complaint analysis - start the web UI
rem
rem  Run from a terminal:   scripts\start_webapp.bat
rem  Or just double-click it.
rem
rem  The window stays open on any error so you can read it.
rem  Override with:  set ECOCO_PORT=8600   /  set ECOCO_PYTHON=C:\...\python.exe
rem
rem  NOTE: keep this file ASCII-only. Windows parses .bat with the
rem  OEM codepage, so UTF-8 Chinese comments break the parser.
rem ============================================================
setlocal EnableDelayedExpansion

chcp 65001 >nul
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"

set "PROJECT_DIR=%~dp0.."
pushd "%PROJECT_DIR%"

if exist "scripts\env.local.bat" call "scripts\env.local.bat"

echo.
echo   ECOCO complaint analysis - web UI
echo   ============================================================

rem ---- 1. find a Python that actually has streamlit -----------
set "PYEXE="
if not "%ECOCO_PYTHON%"=="" (
    call :try_python "%ECOCO_PYTHON%"
)
if "%PYEXE%"=="" call :try_python "%LOCALAPPDATA%\Python\pythoncore-3.14-64\python.exe"
if "%PYEXE%"=="" call :try_python "python"
if "%PYEXE%"=="" call :try_python "py"

if "%PYEXE%"=="" (
    echo.
    echo   [ERROR] No Python with streamlit was found.
    echo.
    echo   Install the dependencies first:
    echo       pip install -r requirements.txt
    echo.
    echo   Or point at the right interpreter:
    echo       set ECOCO_PYTHON=C:\path\to\python.exe
    echo.
    goto :fail
)
echo   Python  : %PYEXE%

rem ---- 2. pick a free port ------------------------------------
if "%ECOCO_PORT%"=="" set "ECOCO_PORT=8501"

rem ---- already running? ---------------------------------------
rem  Scheduled/startup launches must not spawn a second copy on another
rem  port. ECOCO_NO_PORT_HUNT=1 makes a busy port a clean no-op exit.
set "BUSY="
for /f %%b in ('powershell -NoProfile -Command "if(Get-NetTCPConnection -LocalPort %ECOCO_PORT% -State Listen -ErrorAction SilentlyContinue){'1'}else{'0'}"') do set "BUSY=%%b"
if "%BUSY%"=="1" if "%ECOCO_NO_PORT_HUNT%"=="1" (
    echo   Already serving on port %ECOCO_PORT% - nothing to do.
    popd
    endlocal
    exit /b 0
)

set "PORT="
for /f %%p in ('powershell -NoProfile -Command "$s=%ECOCO_PORT%; for($p=$s;$p -lt ($s+20);$p++){ if(-not (Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue)){ $p; break } }"') do set "PORT=%%p"

if "%PORT%"=="" (
    echo.
    echo   [ERROR] No free port between %ECOCO_PORT% and %ECOCO_PORT%+19.
    echo   Close whatever is using them, or pick another range:
    echo       set ECOCO_PORT=8600
    echo.
    goto :fail
)
if not "%PORT%"=="%ECOCO_PORT%" (
    echo   [NOTE]  Port %ECOCO_PORT% was busy, using %PORT% instead.
)

rem ---- 3. show the URLs and launch ----------------------------
set "LANIP="
for /f %%i in ('powershell -NoProfile -Command "foreach($a in @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)){ if($a.IPAddress -notlike '127.*' -and $a.IPAddress -notlike '169.254.*'){ $a.IPAddress; break } }"') do set "LANIP=%%i"

echo   ------------------------------------------------------------
echo   Local   : http://localhost:%PORT%
if not "%LANIP%"=="" echo   Network : http://%LANIP%:%PORT%
echo   ------------------------------------------------------------
echo   Press Ctrl+C to stop. Closing this window also stops it.
echo.

"%PYEXE%" -m streamlit run complaint_webapp.py --server.port %PORT% --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo   [ERROR] Streamlit exited with code %RC%. See the message above.
    goto :fail
)
popd
endlocal
exit /b 0

rem ---- helper: set PYEXE if the candidate has streamlit -------
:try_python
rem  Compare the printed token, not just the exit code: a non-Python exe
rem  (cmd.exe for one) can return 0 and would otherwise be accepted.
if not "%PYEXE%"=="" goto :eof
set "_probe="
for /f "delims=" %%r in ('%~1 -c "import streamlit,sys;sys.stdout.write('ECOCO_OK')" 2^>nul') do set "_probe=%%r"
if "%_probe%"=="ECOCO_OK" set "PYEXE=%~1"
set "_probe="
goto :eof

:fail
echo.
pause
popd
endlocal
exit /b 1
