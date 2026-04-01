@echo off
chcp 65001 >nul
echo === C# Add-Type Compatibility Check ===
echo.
echo Testing if Add-Type (csc.exe) is allowed in this environment...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Check.ps1"
pause
