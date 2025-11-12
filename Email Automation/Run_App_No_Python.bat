@echo off
REM Email Folder Attachment App Launcher
REM This batch file launches the PowerShell GUI application

echo Starting Email Folder Attachment App...
echo.

REM Check if PowerShell is available
powershell -Command "Write-Host 'PowerShell is available'" >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: PowerShell is not available or accessible.
    echo Please contact your IT administrator.
    pause
    exit /b 1
)

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

REM Run the PowerShell script
powershell -ExecutionPolicy Bypass -File "%SCRIPT_DIR%EmailFolderApp.ps1"

REM Check if there was an error
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo An error occurred while running the application.
    echo Error code: %ERRORLEVEL%
    pause
)