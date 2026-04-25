@echo off
setlocal
echo === Demo 2: Environ$ Solution for OneDrive Paths ===
echo.
echo Generating demo xlsm...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Generate-EnvironSolutionDemo.ps1" -OutputDirectory "%~dp0output"
echo.
echo Done. Open the generated xlsm and run Demo_EnvironSolution (Alt+F8).
pause
