@echo off
setlocal
echo === Diff ===
echo Input: %*
pause
if "%~2"=="" (
    echo Usage: Diff.bat file1.xlsm file2.xlsm
    pause
    exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Diff.ps1" "%~1" "%~2"
echo.
echo Done. Exit code: %errorlevel%
pause
