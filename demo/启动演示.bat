@echo off
rem ============================================================
rem  Oncall Dashboard Demo launcher
rem  NOTE: keep this file pure ASCII. cmd.exe reads .bat files
rem        with the OEM code page (936 on zh-CN) and will
rem        corrupt any UTF-8 Chinese text, breaking the parser.
rem ============================================================
setlocal
cd /d "%~dp0"

echo ============================================
echo   Oncall Dashboard Demo  (no Docker needed)
echo ============================================
echo.

python --version >nul 2>nul
if not errorlevel 1 (
    set "PY=python"
    goto run
)

py -3 --version >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
    goto run
)

echo [ERROR] Python not found in PATH.
echo         Install Python 3.8+ first, then run this again.
echo.
pause
exit /b 1

:run
echo Starting mock exporter on port 9100 ...
start "oncall-demo-exporter" %PY% mock_exporter.py --port 9100
timeout /t 3 >nul

echo Opening dashboard preview ...
start "" "preview.html"

echo.
echo Suggested demo steps (click the buttons in order):
echo   1) Normal
echo   2) Latency degrade  -- links stay UP, but P90/P99 rise
echo   3) Link down        -- both links down, connections drop to 0
echo   4) Recover
echo.
echo To stop: close the window titled "oncall-demo-exporter".
echo.
pause
