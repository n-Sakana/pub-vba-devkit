@echo off
setlocal
echo === Demo 1: Path Problem on OneDrive Synced Folders ===
echo.
echo Generating demo xlsm...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Generate-PathProblemDemo.ps1" -OutputDirectory "%~dp0output"
echo.
echo Done. Open the generated xlsm and run Demo_PathProblem (Alt+F8).
pause
