@echo off
setlocal
echo === Unlock ===
echo Input: %*
pause
if "%~1"=="" (
    echo Drop an Excel file to remove VBA project password.
    pause
    exit /b 1
)
for %%F in (%*) do (
    echo.
    echo ----------------------------------------
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Unlock.ps1" "%%~F"
)
echo.
echo Done. Exit code: %errorlevel%
pause
