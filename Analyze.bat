@echo off
setlocal enabledelayedexpansion
set "BATDIR=%~dp0"
echo === Analyze ===
echo Input: %*
pause
if "%~1"=="" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Analyze.ps1"
    pause
    exit /b
)
set "args="
:args_loop
if "%~1"=="" goto :args_done
set "args=!args! "%~1""
shift
goto :args_loop
:args_done
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Analyze.ps1" %args%
echo.
echo Done. Exit code: %errorlevel%
pause
