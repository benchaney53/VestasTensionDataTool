@echo off
REM ========================================
REM Tower Bolt Tension Data Tool - Standalone GUI Installer
REM ========================================
REM This installer downloads from GitHub and sets up the application
REM No Git required - works with PowerShell only

REM Run PowerShell installer completely hidden
PowerShell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -Command "& '%~dp0Install-Tool.ps1'"

REM If there was an error, show a message box instead of terminal output
if %errorLevel% NEQ 0 (
    PowerShell.exe -ExecutionPolicy Bypass -NoProfile -Command "[System.Windows.Forms.MessageBox]::Show('Installation failed. Please check that Python is installed at C:\ProgramData\anaconda3\python.exe and you have an internet connection.', 'Installation Error', 'OK', 'Error')"
)

exit /b %errorLevel%
