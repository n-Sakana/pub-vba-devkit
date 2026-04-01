@echo off
setlocal
echo === Demo 3: Win32 API via HTTP Backend (PS+C#) ===
echo.
echo [1/2] Starting backend server...
start "Backend Server" powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\BackendServer.ps1"
timeout /t 2 >nul
echo [2/2] Generating client xlsm...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0lib\Generate-BackendClientDemo.ps1" -OutputDirectory "%~dp0output"
echo.
echo Done. Open the generated xlsm and run macros (Alt+F8):
echo   - Demo_Win32Api : system info via Win32 API
echo   - Demo_Wiggle   : Excel window wiggle via SetWindowPos
echo.
echo Press any key to shut down the server when done...
pause >nul
echo Shutting down server...
powershell.exe -NoProfile -Command "try { (New-Object Net.WebClient).DownloadString('http://127.0.0.1:8899/api/shutdown') } catch {}" 2>nul
echo Server stopped.
