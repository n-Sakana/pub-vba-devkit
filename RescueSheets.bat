@echo off
setlocal enabledelayedexpansion
set "BATDIR=%~dp0"
echo === RescueSheets ===
echo Input: %*
pause
if "%~1"=="" (
    echo Drop an xlsm/xltm/xlam/xls file or folder to create a macro-disabled sheet rescue copy.
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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\RescueSheets.ps1" %args%
echo.
echo Done. Exit code: %errorlevel%
pause
