@echo off
REM ========================================
REM Tower Bolt Tension Data Tool - Setup Launcher
REM ========================================

title Tower Bolt Tension Data Tool - Setup

echo.
echo ========================================
echo   Tower Bolt Tension Data Tool
echo   Setup Wizard
echo ========================================
echo.

echo Starting setup wizard...
echo.

REM Run PowerShell setup with proper execution policy
PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command "& '%~dp0Setup.exe.ps1'"

if %errorLevel% NEQ 0 (
    echo.
    echo Setup encountered an error.
    echo.
    pause
)

echo.
echo Setup completed.
echo.
pause
exit /b 0
