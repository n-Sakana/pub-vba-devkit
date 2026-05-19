@echo off
setlocal enabledelayedexpansion
set "BATDIR=%~dp0"

echo === Sanitize ===
echo Input: %*
pause

if "%~1"=="" (
    echo Usage: Sanitize.bat [mode 1-10] ^<file-or-folder^> [...]
    pause
    exit /b 1
)

set "mode=1"
set "first=%~1"
if "%first%"=="1" set "mode=1" & shift
if "%first%"=="2" set "mode=2" & shift
if "%first%"=="3" set "mode=3" & shift
if "%first%"=="4" set "mode=4" & shift
if "%first%"=="5" set "mode=5" & shift
if "%first%"=="6" set "mode=6" & shift
if "%first%"=="7" set "mode=7" & shift
if "%first%"=="8" set "mode=8" & shift
if "%first%"=="9" set "mode=9" & shift
if "%first%"=="10" set "mode=10" & shift

if "%~1"=="" (
    echo Usage: Sanitize.bat [mode 1-10] ^<file-or-folder^> [...]
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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%lib\Sanitize.ps1" -Mode %mode% -Path %args%
echo.
echo Done. Exit code: %errorlevel%
pause
