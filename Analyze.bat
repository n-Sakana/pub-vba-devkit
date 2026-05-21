@echo off
setlocal
set "BATDIR=%~dp0"
echo === Analyze ===
echo Input: %*
pause
if "%~1"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Analyze.ps1"
    pause
    exit /b
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Analyze.ps1" %*
echo.
echo Done. Exit code: %errorlevel%
pause
