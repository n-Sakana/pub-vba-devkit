@echo off
setlocal
set "BATDIR=%~dp0"
echo === Extract ===
echo Input: %*
pause
if "%~1"=="" (
    echo Drop Excel files or folder to extract VBA code.
    pause
    exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Extract.ps1" %*
echo.
echo Done. Exit code: %errorlevel%
pause
