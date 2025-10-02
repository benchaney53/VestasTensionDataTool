@echo off
setlocal enabledelayedexpansion
REM ========================================
REM Tower Bolt Tension Data Tool
REM Standalone Installer - Downloads Everything from GitHub
REM ========================================

title Tower Bolt Tension Data Tool - Installer

echo.
echo ========================================
echo   Tower Bolt Tension Data Tool
echo   Installation Wizard
echo ========================================
echo.

REM Configuration
set "GITHUB_USER=benchaney53"
set "REPO_NAME=VestasTensionDataTool"
set "BRANCH=master"
set "APP_NAME=Tower Bolt Tension Data Tool"

REM Check for Python in multiple locations
echo Checking for Python installation...

set "PYTHON_EXE="

REM Try python in PATH first
python --version >nul 2>&1
if %errorLevel% EQU 0 (
    set "PYTHON_EXE=python"
    for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
    echo [OK] Found in PATH: %PYTHON_VERSION%
    goto :python_found
)

REM Check common Python installation locations
set "PYTHON_PATHS=C:\Python39\python.exe;C:\Python310\python.exe;C:\Python311\python.exe;C:\Python312\python.exe;C:\Python313\python.exe;%LOCALAPPDATA%\Programs\Python\Python39\python.exe;%LOCALAPPDATA%\Programs\Python\Python310\python.exe;%LOCALAPPDATA%\Programs\Python\Python311\python.exe;%LOCALAPPDATA%\Programs\Python\Python312\python.exe;%LOCALAPPDATA%\Programs\Python\Python313\python.exe;C:\ProgramData\anaconda3\python.exe;%USERPROFILE%\anaconda3\python.exe;%USERPROFILE%\miniconda3\python.exe;C:\ProgramData\miniconda3\python.exe"

for %%p in ("%PYTHON_PATHS:;=" "%") do (
    if exist %%p (
        "%%~p" --version >nul 2>&1
        if !errorLevel! EQU 0 (
            set "PYTHON_EXE=%%~p"
            for /f "tokens=*" %%i in ('"%%~p" --version') do set PYTHON_VERSION=%%i
            echo [OK] Found: !PYTHON_VERSION! at %%~p
            goto :python_found
        )
    )
)

REM If no Python found
echo.
echo [ERROR] Python installation not found!
echo.
echo Searched locations:
echo   - System PATH
echo   - C:\Python3x\
echo   - %LOCALAPPDATA%\Programs\Python\
echo   - Anaconda/Miniconda installations
echo.
echo Please install Python 3.8 or later from:
echo https://www.python.org/downloads/
echo.
echo Make sure to check "Add Python to PATH" during installation.
echo.
pause
exit /b 1

:python_found
echo.

REM Ask for installation location
set "DEFAULT_PATH=%USERPROFILE%\Desktop\%APP_NAME%"
echo Where would you like to install %APP_NAME%?
echo Default: %DEFAULT_PATH%
echo.
set /p "INSTALL_PATH=Press Enter for default or type a custom path: "

if "%INSTALL_PATH%"=="" set "INSTALL_PATH=%DEFAULT_PATH%"

echo.
echo Installation path: %INSTALL_PATH%
echo.

REM Check if directory exists
if exist "%INSTALL_PATH%" (
    echo WARNING: Directory already exists. Files will be overwritten.
    set /p "CONTINUE=Continue? (Y/N): "
    if /i not "%CONTINUE%"=="Y" (
        echo Installation cancelled.
        pause
        exit /b 0
    )
) else (
    echo Creating installation directory...
    mkdir "%INSTALL_PATH%" 2>nul
    if %errorLevel% NEQ 0 (
        echo [ERROR] Failed to create directory. Check permissions.
        pause
        exit /b 1
    )
)

echo.
echo ========================================
echo Downloading from GitHub...
echo ========================================
echo.

REM Download ZIP from GitHub
set "ZIP_URL=https://github.com/%GITHUB_USER%/%REPO_NAME%/archive/refs/heads/%BRANCH%.zip"
set "TEMP_ZIP=%TEMP%\%REPO_NAME%.zip"
set "TEMP_EXTRACT=%TEMP%\%REPO_NAME%-extract"

echo Downloading: %ZIP_URL%
echo.

PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
    "  Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile '%TEMP_ZIP%' -UseBasicParsing; " ^
    "  Write-Host '[OK] Download complete' -ForegroundColor Green; " ^
    "} catch { " ^
    "  Write-Host '[ERROR] Download failed' -ForegroundColor Red; " ^
    "  Write-Host $_.Exception.Message -ForegroundColor Red; " ^
    "  exit 1; " ^
    "}"

if %errorLevel% NEQ 0 (
    echo.
    echo Installation failed.
    pause
    exit /b 1
)

echo.
echo Extracting files...

REM Clean up old extraction folder
if exist "%TEMP_EXTRACT%" rmdir /s /q "%TEMP_EXTRACT%"

PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
    "try { " ^
    "  Expand-Archive -Path '%TEMP_ZIP%' -DestinationPath '%TEMP_EXTRACT%' -Force; " ^
    "  $extractedFolder = Get-ChildItem '%TEMP_EXTRACT%' | Select-Object -First 1; " ^
    "  Copy-Item \"$($extractedFolder.FullName)\*\" -Destination '%INSTALL_PATH%' -Recurse -Force; " ^
    "  Write-Host '[OK] Files extracted' -ForegroundColor Green; " ^
    "} catch { " ^
    "  Write-Host '[ERROR] Extraction failed' -ForegroundColor Red; " ^
    "  Write-Host $_.Exception.Message -ForegroundColor Red; " ^
    "  exit 1; " ^
    "}"

if %errorLevel% NEQ 0 (
    echo.
    echo Installation failed.
    pause
    exit /b 1
)

REM Cleanup temp files
del "%TEMP_ZIP%" 2>nul
rmdir /s /q "%TEMP_EXTRACT%" 2>nul

echo.
echo ========================================
echo Setting up Python environment...
echo ========================================
echo.

set "VENV_PATH=%INSTALL_PATH%\venv"

REM Remove old venv if exists
if exist "%VENV_PATH%" (
    echo Removing old virtual environment...
    rmdir /s /q "%VENV_PATH%"
)

echo Creating virtual environment...
echo Using Python: %PYTHON_EXE%
"%PYTHON_EXE%" -m venv "%VENV_PATH%"

if %errorLevel% NEQ 0 (
    echo [ERROR] Failed to create virtual environment
    echo Python executable: %PYTHON_EXE%
    pause
    exit /b 1
)

echo [OK] Virtual environment created
echo.

REM Install dependencies
set "REQUIREMENTS=%INSTALL_PATH%\app\requirements.txt"

if exist "%REQUIREMENTS%" (
    echo Installing dependencies ^(this may take a few minutes^)...
    echo Please wait...
    echo.
    
    "%VENV_PATH%\Scripts\python.exe" -m pip install --upgrade pip --quiet
    "%VENV_PATH%\Scripts\pip.exe" install -r "%REQUIREMENTS%" --quiet
    
    if %errorLevel% EQU 0 (
        echo [OK] Dependencies installed
    ) else (
        echo [WARNING] Some dependencies may have failed to install
    )
) else (
    echo [WARNING] requirements.txt not found
)

echo.
echo ========================================
echo Creating launcher...
echo ========================================
echo.

REM Create launcher batch file
(
echo @echo off
echo REM Tower Bolt Tension Data Tool Launcher
echo REM Created by installer using: %PYTHON_EXE%
echo cd /d "%INSTALL_PATH%"
echo "%VENV_PATH%\Scripts\python.exe" "%INSTALL_PATH%\app\main.py"
echo if %%errorLevel%% NEQ 0 ^(
echo     echo.
echo     echo [ERROR] Application failed to start
echo     echo Check that all dependencies are installed correctly
echo     echo.
echo ^)
echo pause
) > "%INSTALL_PATH%\Launch.bat"

echo [OK] Launcher created

REM Create Python path info file for troubleshooting
(
echo Tower Bolt Tension Data Tool - Installation Info
echo ================================================
echo.
echo Installation Date: %DATE% %TIME%
echo Python Executable Used: %PYTHON_EXE%
echo Python Version: %PYTHON_VERSION%
echo Installation Path: %INSTALL_PATH%
echo Virtual Environment: %VENV_PATH%
echo.
echo This file is for troubleshooting purposes only.
) > "%INSTALL_PATH%\installation_info.txt"

echo.

REM Ask about desktop shortcut
set /p "CREATE_SHORTCUT=Create desktop shortcut? (Y/N): "

if /i "%CREATE_SHORTCUT%"=="Y" (
    PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^
        "try { " ^
        "  $WshShell = New-Object -ComObject WScript.Shell; " ^
        "  $Shortcut = $WshShell.CreateShortcut('%USERPROFILE%\Desktop\%APP_NAME%.lnk'); " ^
        "  $Shortcut.TargetPath = '%INSTALL_PATH%\Launch.bat'; " ^
        "  $Shortcut.WorkingDirectory = '%INSTALL_PATH%'; " ^
        "  $Shortcut.Description = '%APP_NAME%'; " ^
        "  $Shortcut.Save(); " ^
        "  Write-Host '[OK] Desktop shortcut created' -ForegroundColor Green; " ^
        "} catch { " ^
        "  Write-Host '[WARNING] Failed to create shortcut' -ForegroundColor Yellow; " ^
        "}"
)

echo.
echo ========================================
echo   Installation Complete!
echo ========================================
echo.
echo Installation Location:
echo   %INSTALL_PATH%
echo.
echo To run the application:
echo   - Use the desktop shortcut ^(if created^)
echo   - Run: %INSTALL_PATH%\Launch.bat
echo   - Or double-click Launch.bat in the installation folder
echo.

set /p "RUN_NOW=Would you like to launch the application now? (Y/N): "

if /i "%RUN_NOW%"=="Y" (
    echo.
    echo Launching application...
    start "" "%INSTALL_PATH%\Launch.bat"
)

echo.
echo Thank you for installing %APP_NAME%!
echo.
pause
exit /b 0

