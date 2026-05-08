@echo off
setlocal

set CHROME=C:\Users\hai.a.nguyen\AppData\Local\Google\Chrome\Application\chrome.exe
set PORT=8888
set URL=http://localhost:%PORT%/popup.html
set SCRIPT=%~dp0server.ps1

echo.
echo  ============================================================
echo   D365 Entity Compare - Local Server Launcher
echo  ============================================================
echo.
echo  Starting HTTP server on %URL%
echo  Close this window to stop the server.
echo.

:: Start the PS server in a separate window (stays open)
start "D365 Server - close to stop" powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"

:: Wait for server to start
timeout /t 2 /nobreak >nul

:: Open Chrome
start "" "%CHROME%" "%URL%"

echo  Chrome opened. Server is running in the other window.
echo  Log in to your D365 environments in Chrome, then use the tool.
echo.
pause
