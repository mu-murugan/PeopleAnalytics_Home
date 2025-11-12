@echo off
title Email Folder Attachment App
echo Starting Email Folder Attachment App...
echo.
echo Current user: %USERNAME%
echo Computer: %COMPUTERNAME%
echo Domain: %USERDOMAIN%
echo.

cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python from https://python.org
    echo.
    pause
    exit /b 1
)

echo Starting application...
python email_folder_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred while running the application.
    echo Error code: %ERRORLEVEL%
    echo.
    echo Press any key to exit.
    pause >nul
)