@echo off
title BP Generator

echo Starting BP Generator...
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Run INSTALLER.bat first.
    pause
    exit /b 1
)

for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| find ":8765 "') do (
    taskkill /PID %%a /F >nul 2>&1
)

start /min "BP-Server" cmd /c "cd /d "%~dp0backend" && python -m uvicorn server:app --host 127.0.0.1 --port 8765"

echo Waiting for server to start...
timeout /t 5 /nobreak >nul

start "" "http://127.0.0.1:8765"

echo.
echo ====================================================
echo  BP Generator is running!
echo  Browser opened on http://127.0.0.1:8765
echo  Access code: JRC-MATEO-2025
echo.
echo  Close this window to stop the application.
echo ====================================================
echo.
pause >nul

for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| find ":8765 "') do (
    taskkill /PID %%a /F >nul 2>&1
)
