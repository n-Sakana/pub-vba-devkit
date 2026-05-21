@echo off
setlocal
set "BATDIR=%~dp0"

echo === Sanitize ===
echo Input: %*
pause

if "%~1"=="" (
    echo Usage: Sanitize.bat ^<file-or-folder^> [...]
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Sanitize.ps1" -Path %*
echo.
echo Done. Exit code: %errorlevel%
pause
