@echo off
REM ========================================
REM Tower Bolt Tension Data Tool - Installer
REM ========================================

echo.
echo ========================================
echo   Tower Bolt Tension Data Tool
echo   Installation Wizard
echo ========================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
    echo.
) else (
    echo Note: Running without administrator privileges.
    echo Some features may require admin rights.
    echo.
)

echo Starting installer...
echo.

REM Run PowerShell installer
PowerShell.exe -ExecutionPolicy Bypass -File "%~dp0installer.ps1"

if %errorLevel% NEQ 0 (
    echo.
    echo Installation failed or was cancelled.
    pause
    exit /b %errorLevel%
)

exit /b 0

