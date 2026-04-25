@echo off
setlocal
set "BATDIR=%~dp0"
echo === Rebuild Test xlsm Files ===
echo.
echo This requires Excel with "Trust access to the VBA project object model" enabled.
echo.
pause
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%BATDIR%rebuild_tests.ps1"
echo.
echo Done. Exit code: %errorlevel%
pause
