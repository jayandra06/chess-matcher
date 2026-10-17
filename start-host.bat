@echo off
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js is required to host this page. Install it from https://nodejs.org and try again.
  pause
  exit /b 1
)
node server.js
echo.
echo The host stopped.
pause
