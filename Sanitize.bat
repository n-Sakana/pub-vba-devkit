@echo off
setlocal enabledelayedexpansion
set "BATDIR=%~dp0"
echo === Sanitize ===
echo Input: %*
pause
if "%~1"=="" (
    echo Drop an xlsm file to sanitize EDR-triggering VBA patterns.
    pause
    exit /b 1
)
set "args="
:args_loop
if "%~1"=="" goto :args_done
set "args=!args! "%~1""
shift
goto :args_loop
:args_done
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Sanitize.ps1" -Path %args%
echo.
echo Done. Exit code: %errorlevel%
pause
