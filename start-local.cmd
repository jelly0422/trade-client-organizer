@echo off
setlocal
cd /d "%~dp0"
set "TRADE_NODE=%~dp0runtime\node-v24.21.0-win-x64\node.exe"
if not exist "%TRADE_NODE%" (
  echo Local Node.js runtime is missing. See LOCAL-SETUP.md.
  pause
  exit /b 1
)
echo Trade Client Organizer
echo Open http://127.0.0.1:4173 in your browser after the server starts.
echo Keep this window open while using the app. Press Ctrl+C to stop.
echo.
"%TRADE_NODE%" "%~dp0server.cjs"
if errorlevel 1 (
  echo.
  echo Server failed to start. Port 4173 may already be in use.
  pause
  exit /b 1
)
