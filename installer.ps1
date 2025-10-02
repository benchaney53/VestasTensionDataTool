# ========================================
# Tower Bolt Tension Data Tool - Installer
# ========================================
# This script will:
# 1. Download the latest version from GitHub
# 2. Set up Python virtual environment
# 3. Install all dependencies
# 4. Create desktop shortcuts
# ========================================

$ErrorActionPreference = "Stop"

# Configuration
$GitHubUser = "benchaney53"
$RepoName = "VestasTensionDataTool"
$Branch = "master"
$AppName = "Tower Bolt Tension Data Tool"

# Colors for output
function Write-Success { Write-Host $args -ForegroundColor Green }
function Write-Info { Write-Host $args -ForegroundColor Cyan }
function Write-Warning { Write-Host $args -ForegroundColor Yellow }
function Write-Error { Write-Host $args -ForegroundColor Red }

# Banner
Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  $AppName" -ForegroundColor Cyan
Write-Host "  Installation Wizard" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for Python
Write-Info "Checking for Python installation..."
try {
    $pythonVersion = python --version 2>&1
    Write-Success "✓ Found: $pythonVersion"
} catch {
    Write-Error "✗ Python is not installed or not in PATH"
    Write-Host ""
    Write-Host "Please install Python 3.8 or later from https://www.python.org/downloads/"
    Write-Host "Make sure to check 'Add Python to PATH' during installation."
    Read-Host "Press Enter to exit"
    exit 1
}

# Ask for installation directory
Write-Host ""
$defaultPath = "$env:USERPROFILE\Desktop\$AppName"
Write-Info "Where would you like to install $AppName?"
Write-Host "Default: $defaultPath"
$installPath = Read-Host "Press Enter for default or type a custom path"

if ([string]::IsNullOrWhiteSpace($installPath)) {
    $installPath = $defaultPath
}

# Create installation directory
Write-Host ""
Write-Info "Creating installation directory..."
if (Test-Path $installPath) {
    Write-Warning "Directory already exists. Existing files will be overwritten."
    $response = Read-Host "Continue? (Y/N)"
    if ($response -ne "Y" -and $response -ne "y") {
        Write-Host "Installation cancelled."
        Read-Host "Press Enter to exit"
        exit 0
    }
} else {
    New-Item -ItemType Directory -Path $installPath -Force | Out-Null
}

# Download from GitHub
Write-Host ""
Write-Info "Downloading latest version from GitHub..."
$zipUrl = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
$zipPath = "$env:TEMP\$RepoName.zip"
$extractPath = "$env:TEMP\$RepoName-extract"

try {
    # Download
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Write-Success "✓ Download complete"
    
    # Extract
    Write-Info "Extracting files..."
    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
    
    # Copy files to installation directory
    $extractedFolder = Get-ChildItem $extractPath | Select-Object -First 1
    Copy-Item "$($extractedFolder.FullName)\*" -Destination $installPath -Recurse -Force
    
    Write-Success "✓ Files extracted"
    
    # Cleanup
    Remove-Item $zipPath -Force
    Remove-Item $extractPath -Recurse -Force
    
} catch {
    Write-Error "✗ Failed to download or extract files"
    Write-Error $_.Exception.Message
    Read-Host "Press Enter to exit"
    exit 1
}

# Create virtual environment
Write-Host ""
Write-Info "Setting up Python virtual environment..."
$venvPath = "$installPath\venv"

try {
    if (Test-Path $venvPath) {
        Remove-Item $venvPath -Recurse -Force
    }
    
    python -m venv $venvPath
    Write-Success "✓ Virtual environment created"
} catch {
    Write-Error "✗ Failed to create virtual environment"
    Write-Error $_.Exception.Message
    Read-Host "Press Enter to exit"
    exit 1
}

# Install dependencies
Write-Host ""
Write-Info "Installing dependencies (this may take a few minutes)..."
$requirementsPath = "$installPath\app\requirements.txt"

if (Test-Path $requirementsPath) {
    try {
        & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip | Out-Null
        & "$venvPath\Scripts\pip.exe" install -r $requirementsPath
        Write-Success "✓ Dependencies installed"
    } catch {
        Write-Error "✗ Failed to install dependencies"
        Write-Error $_.Exception.Message
        Read-Host "Press Enter to exit"
        exit 1
    }
} else {
    Write-Warning "⚠ requirements.txt not found, skipping dependency installation"
}

# Create launcher script
Write-Host ""
Write-Info "Creating launcher scripts..."

# PowerShell launcher
$launcherPS = @"
# Tower Bolt Tension Data Tool Launcher
Set-Location "$installPath"
& "$venvPath\Scripts\python.exe" "$installPath\app\main.py"
"@
$launcherPS | Out-File -FilePath "$installPath\Launch.ps1" -Encoding UTF8

# Batch launcher
$launcherBat = @"
@echo off
cd /d "$installPath"
"$venvPath\Scripts\python.exe" "$installPath\app\main.py"
pause
"@
$launcherBat | Out-File -FilePath "$installPath\Launch.bat" -Encoding ASCII

Write-Success "✓ Launcher scripts created"

# Create desktop shortcut
Write-Host ""
$createShortcut = Read-Host "Create desktop shortcut? (Y/N)"
if ($createShortcut -eq "Y" -or $createShortcut -eq "y") {
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$AppName.lnk")
        $Shortcut.TargetPath = "$installPath\Launch.bat"
        $Shortcut.WorkingDirectory = $installPath
        $Shortcut.Description = "$AppName"
        $Shortcut.Save()
        Write-Success "✓ Desktop shortcut created"
    } catch {
        Write-Warning "⚠ Failed to create desktop shortcut"
    }
}

# Installation complete
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Success "  Installation Complete!"
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Info "Installation Location:"
Write-Host "  $installPath" -ForegroundColor White
Write-Host ""
Write-Info "To run the application:"
Write-Host "  - Use the desktop shortcut (if created)" -ForegroundColor White
Write-Host "  - Run: $installPath\Launch.bat" -ForegroundColor White
Write-Host "  - Or double-click: Launch.bat in the installation folder" -ForegroundColor White
Write-Host ""

$runNow = Read-Host "Would you like to launch the application now? (Y/N)"
if ($runNow -eq "Y" -or $runNow -eq "y") {
    Write-Info "Launching application..."
    Start-Process -FilePath "$installPath\Launch.bat"
}

Write-Host ""
Write-Success "Thank you for installing $AppName!"
Read-Host "Press Enter to exit"

