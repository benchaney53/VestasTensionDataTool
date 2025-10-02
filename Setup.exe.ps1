# ========================================
# Tower Bolt Tension Data Tool - Setup Wizard
# ========================================
# Professional GUI Installer for Windows
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Configuration
$GitHubUser = "benchaney53"
$RepoName = "VestasTensionDataTool"
$Branch = "master"
$AppName = "Tower Bolt Tension Data Tool"
$AppVersion = "1.0.0"

# Global variables
$script:InstallPath = ""
$script:PythonPath = ""
$script:CurrentStep = 0
$script:TotalSteps = 5
$script:Wizard = $null
$script:ProgressBar = $null
$script:StatusLabel = $null
$script:NextButton = $null
$script:BackButton = $null
$script:CancelButton = $null

# Find Python installation
function Find-Python {
    $pythonPaths = @(
        "python",  # In PATH
        "$env:ProgramData\anaconda3\python.exe",
        "$env:USERPROFILE\anaconda3\python.exe",
        "$env:USERPROFILE\miniconda3\python.exe",
        "$env:ProgramData\miniconda3\python.exe",
        "C:\Python39\python.exe",
        "C:\Python310\python.exe",
        "C:\Python311\python.exe",
        "C:\Python312\python.exe",
        "C:\Python313\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python39\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python313\python.exe"
    )

    foreach ($path in $pythonPaths) {
        try {
            $version = & $path "--version" 2>$null
            if ($LASTEXITCODE -eq 0 -and $version -match "Python (\d+)\.(\d+)\.(\d+)") {
                $major = [int]$Matches[1]
                $minor = [int]$Matches[2]
                if ($major -eq 3 -and $minor -ge 8) {
                    return @{
                        Path = $path
                        Version = $version.Trim()
                        DisplayName = if ($path -eq "python") { "Python (in PATH)" } else { "Python at $path" }
                    }
                }
            }
        } catch {
            continue
        }
    }
    return $null
}

# Update progress
function Update-Progress {
    param($step, $message, $detail = "")
    $script:CurrentStep = $step
    $script:ProgressBar.Value = ($step / $script:TotalSteps) * 100
    $script:StatusLabel.Text = $message
    if ($detail) {
        $script:StatusLabel.Text += "`n$detail"
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Download file with progress
function Download-File {
    param($url, $outputPath)

    try {
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $outputPath)
        return $true
    } catch {
        return $false
    }
}

# Create the main wizard form
function Create-WizardForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$AppName Setup Wizard"
    $form.Size = New-Object System.Drawing.Size(600, 450)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Icon = [System.Drawing.SystemIcons]::Application
    $form.BackColor = [System.Drawing.Color]::White

    return $form
}

# Welcome page
function Show-WelcomePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    $panel.BackColor = [System.Drawing.Color]::White

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Welcome to $AppName Setup"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 114, 198)  # Vestas blue
    $panel.Controls.Add($titleLabel)

    # Subtitle
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "Version $AppVersion"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 60)
    $subtitleLabel.Size = New-Object System.Drawing.Size(550, 20)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::Gray
    $panel.Controls.Add($subtitleLabel)

    # Description
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "This wizard will install $AppName on your computer.`n`n" +
                     "The setup will:`n" +
                     "• Download the latest version from GitHub`n" +
                     "• Set up a Python virtual environment`n" +
                     "• Install all required dependencies`n" +
                     "• Create shortcuts for easy access`n`n" +
                     "Click Next to continue."
    $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $descLabel.Location = New-Object System.Drawing.Point(20, 90)
    $descLabel.Size = New-Object System.Drawing.Size(550, 120)
    $panel.Controls.Add($descLabel)

    # Vestas logo area (placeholder)
    $logoLabel = New-Object System.Windows.Forms.Label
    $logoLabel.Text = "VESTAS"
    $logoLabel.Font = New-Object System.Drawing.Font("Arial", 24, [System.Drawing.FontStyle]::Bold)
    $logoLabel.Location = New-Object System.Drawing.Point(20, 220)
    $logoLabel.Size = New-Object System.Drawing.Size(200, 50)
    $logoLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 114, 198)
    $panel.Controls.Add($logoLabel)

    return $panel
}

# License/Terms page
function Show-LicensePage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    $panel.BackColor = [System.Drawing.Color]::White

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "License Agreement"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 30)
    $panel.Controls.Add($titleLabel)

    $licenseText = New-Object System.Windows.Forms.TextBox
    $licenseText.Multiline = $true
    $licenseText.ReadOnly = $true
    $licenseText.ScrollBars = "Vertical"
    $licenseText.Text = "END-USER LICENSE AGREEMENT FOR TOWER BOLT TENSION DATA TOOL

This End-User License Agreement (""EULA"") is a legal agreement between you and Vestas Wind Systems.

By installing, copying, or otherwise using this software, you agree to be bound by the terms of this EULA.

1. GRANT OF LICENSE
Vestas grants you a non-exclusive, non-transferable license to use this software for internal business purposes only.

2. RESTRICTIONS
You may not reverse engineer, decompile, or disassemble this software.

3. INTELLECTUAL PROPERTY
This software contains proprietary and confidential information of Vestas.

4. LIMITATION OF LIABILITY
In no event shall Vestas be liable for any damages arising from the use of this software.

By clicking ""I Agree"", you acknowledge that you have read this agreement and agree to its terms."
    $licenseText.Location = New-Object System.Drawing.Point(20, 60)
    $licenseText.Size = New-Object System.Drawing.Size(550, 250)
    $licenseText.Font = New-Object System.Drawing.Font("Consolas", 8)
    $panel.Controls.Add($licenseText)

    $agreeCheckBox = New-Object System.Windows.Forms.CheckBox
    $agreeCheckBox.Text = "I agree to the terms of this license agreement"
    $agreeCheckBox.Location = New-Object System.Drawing.Point(20, 320)
    $agreeCheckBox.Size = New-Object System.Drawing.Size(400, 20)
    $agreeCheckBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $panel.Controls.Add($agreeCheckBox)

    # Store checkbox reference for validation
    $script:AgreeCheckBox = $agreeCheckBox

    return $panel
}

# Installation location page
function Show-LocationPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    $panel.BackColor = [System.Drawing.Color]::White

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Choose Installation Location"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 30)
    $panel.Controls.Add($titleLabel)

    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "Select the folder where $AppName will be installed:"
    $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $descLabel.Location = New-Object System.Drawing.Point(20, 50)
    $descLabel.Size = New-Object System.Drawing.Size(550, 20)
    $panel.Controls.Add($descLabel)

    $defaultPath = "$env:USERPROFILE\Desktop\$AppName"
    $script:InstallPath = $defaultPath

    $pathTextBox = New-Object System.Windows.Forms.TextBox
    $pathTextBox.Text = $defaultPath
    $pathTextBox.Location = New-Object System.Drawing.Point(20, 80)
    $pathTextBox.Size = New-Object System.Drawing.Size(450, 25)
    $pathTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $panel.Controls.Add($pathTextBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse..."
    $browseButton.Location = New-Object System.Drawing.Point(480, 78)
    $browseButton.Size = New-Object System.Drawing.Size(80, 27)
    $browseButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $browseButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select installation folder"
        $folderBrowser.SelectedPath = $pathTextBox.Text
        if ($folderBrowser.ShowDialog() -eq "OK") {
            $pathTextBox.Text = $folderBrowser.SelectedPath
            $script:InstallPath = $folderBrowser.SelectedPath
        }
    })
    $panel.Controls.Add($browseButton)

    $pathTextBox.Add_TextChanged({
        $script:InstallPath = $pathTextBox.Text
    })

    $spaceLabel = New-Object System.Windows.Forms.Label
    $spaceLabel.Text = "Required space: ~500 MB"
    $spaceLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $spaceLabel.ForeColor = [System.Drawing.Color]::Gray
    $spaceLabel.Location = New-Object System.Drawing.Point(20, 120)
    $spaceLabel.Size = New-Object System.Drawing.Size(300, 20)
    $panel.Controls.Add($spaceLabel)

    return $panel
}

# Installation progress page
function Show-InstallPage {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    $panel.BackColor = [System.Drawing.Color]::White

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Installing $AppName"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(550, 30)
    $panel.Controls.Add($titleLabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Preparing installation..."
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $statusLabel.Location = New-Object System.Drawing.Point(20, 60)
    $statusLabel.Size = New-Object System.Drawing.Size(550, 40)
    $panel.Controls.Add($statusLabel)
    $script:StatusLabel = $statusLabel

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 110)
    $progressBar.Size = New-Object System.Drawing.Size(550, 25)
    $progressBar.Style = "Continuous"
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $panel.Controls.Add($progressBar)
    $script:ProgressBar = $progressBar

    $detailLabel = New-Object System.Windows.Forms.Label
    $detailLabel.Text = ""
    $detailLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $detailLabel.ForeColor = [System.Drawing.Color]::Gray
    $detailLabel.Location = New-Object System.Drawing.Point(20, 145)
    $detailLabel.Size = New-Object System.Drawing.Size(550, 60)
    $panel.Controls.Add($detailLabel)

    return $panel
}

# Completion page
function Show-CompletePage {
    param($success = $true, $errorMessage = "")

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"
    $panel.BackColor = [System.Drawing.Color]::White

    if ($success) {
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Installation Complete!"
        $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
        $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(550, 40)
        $titleLabel.ForeColor = [System.Drawing.Color]::Green
        $panel.Controls.Add($titleLabel)

        $successIcon = New-Object System.Windows.Forms.Label
        $successIcon.Text = "✓"
        $successIcon.Font = New-Object System.Drawing.Font("Segoe UI", 48)
        $successIcon.Location = New-Object System.Drawing.Point(500, 10)
        $successIcon.Size = New-Object System.Drawing.Size(60, 60)
        $successIcon.ForeColor = [System.Drawing.Color]::Green
        $panel.Controls.Add($successIcon)

        $descLabel = New-Object System.Windows.Forms.Label
        $descLabel.Text = "$AppName has been successfully installed!`n`n" +
                         "Installation location: $script:InstallPath`n`n" +
                         "You can now launch the application from the desktop shortcut or by running Launch.bat."
        $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $descLabel.Location = New-Object System.Drawing.Point(20, 70)
        $descLabel.Size = New-Object System.Drawing.Size(550, 100)
        $panel.Controls.Add($descLabel)

        $launchCheckBox = New-Object System.Windows.Forms.CheckBox
        $launchCheckBox.Text = "Launch $AppName now"
        $launchCheckBox.Location = New-Object System.Drawing.Point(20, 180)
        $launchCheckBox.Size = New-Object System.Drawing.Size(200, 20)
        $launchCheckBox.Checked = $true
        $panel.Controls.Add($launchCheckBox)
        $script:LaunchCheckBox = $launchCheckBox

    } else {
        $titleLabel = New-Object System.Windows.Forms.Label
        $titleLabel.Text = "Installation Failed"
        $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
        $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
        $titleLabel.Size = New-Object System.Drawing.Size(550, 40)
        $titleLabel.ForeColor = [System.Drawing.Color]::Red
        $panel.Controls.Add($titleLabel)

        $errorIcon = New-Object System.Windows.Forms.Label
        $errorIcon.Text = "✗"
        $errorIcon.Font = New-Object System.Drawing.Font("Segoe UI", 48)
        $errorIcon.Location = New-Object System.Drawing.Point(500, 10)
        $errorIcon.Size = New-Object System.Drawing.Size(60, 60)
        $errorIcon.ForeColor = [System.Drawing.Color]::Red
        $panel.Controls.Add($errorIcon)

        $descLabel = New-Object System.Windows.Forms.Label
        $descLabel.Text = "The installation encountered an error:`n`n$errorMessage`n`n" +
                         "Please check the requirements and try again."
        $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $descLabel.Location = New-Object System.Drawing.Point(20, 70)
        $descLabel.Size = New-Object System.Drawing.Size(550, 100)
        $panel.Controls.Add($descLabel)
    }

    return $panel
}

# Perform the actual installation
function Install-Application {
    try {
        Update-Progress 1 "Checking Python installation..." "Searching for compatible Python version..."

        # Find Python
        $pythonInfo = Find-Python
        if (-not $pythonInfo) {
            throw "Python 3.8 or later is required but not found. Please install Python from https://python.org"
        }
        $script:PythonPath = $pythonInfo.Path
        Update-Progress 1 "Found Python: $($pythonInfo.Version)" "Using: $($pythonInfo.DisplayName)"

        # Create installation directory
        Update-Progress 2 "Creating installation directory..." $script:InstallPath
        if (Test-Path $script:InstallPath) {
            Remove-Item $script:InstallPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $script:InstallPath -Force | Out-Null

        # Download from GitHub
        Update-Progress 3 "Downloading from GitHub..." "Fetching latest version..."
        $zipUrl = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
        $tempZip = "$env:TEMP\$RepoName.zip"
        $tempExtract = "$env:TEMP\$RepoName-extract"

        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }

        Download-File $zipUrl $tempZip
        if (-not (Test-Path $tempZip)) {
            throw "Failed to download from GitHub"
        }

        # Extract
        Update-Progress 3 "Extracting files..." "Unpacking application files..."
        Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
        $extractedFolder = Get-ChildItem $tempExtract | Select-Object -First 1
        Copy-Item "$($extractedFolder.FullName)\*" -Destination $script:InstallPath -Recurse -Force

        # Cleanup
        Remove-Item $tempZip -Force
        Remove-Item $tempExtract -Recurse -Force

        # Create virtual environment
        Update-Progress 4 "Setting up Python environment..." "Creating virtual environment..."
        $venvPath = "$script:InstallPath\venv"
        & $script:PythonPath -m venv $venvPath
        if ($LASTEXITCODE -ne 0) { throw "Failed to create virtual environment" }

        # Install dependencies
        Update-Progress 4 "Installing dependencies..." "This may take a few minutes..."
        $requirementsPath = "$script:InstallPath\app\requirements.txt"
        if (Test-Path $requirementsPath) {
            & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
            & "$venvPath\Scripts\pip.exe" install -r $requirementsPath --quiet
            if ($LASTEXITCODE -ne 0) { throw "Failed to install dependencies" }
        }

        # Create launcher
        Update-Progress 5 "Creating shortcuts..." "Setting up application launcher..."
        $launcherContent = @"
@echo off
REM $AppName Launcher
REM Created by installer using: $script:PythonPath
cd /d "$script:InstallPath"
"$venvPath\Scripts\python.exe" "$script:InstallPath\app\main.py"
if %errorLevel% NEQ 0 (
    echo.
    echo [ERROR] Application failed to start
    echo Check that all dependencies are installed correctly
    echo.
)
pause
"@
        $launcherContent | Out-File -FilePath "$script:InstallPath\Launch.bat" -Encoding UTF8

        # Create installation info
        $infoContent = @"
$AppName - Installation Info
================================
Installation Date: $(Get-Date)
Python Executable: $script:PythonPath
Python Version: $($pythonInfo.Version)
Installation Path: $script:InstallPath
Virtual Environment: $venvPath

This file is for troubleshooting purposes only.
"@
        $infoContent | Out-File -FilePath "$script:InstallPath\installation_info.txt" -Encoding UTF8

        # Create desktop shortcut
        Update-Progress 5 "Creating desktop shortcut..." "Adding to desktop..."
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$AppName.lnk")
        $Shortcut.TargetPath = "$script:InstallPath\Launch.bat"
        $Shortcut.WorkingDirectory = $script:InstallPath
        $Shortcut.Description = $AppName
        $Shortcut.Save()

        Update-Progress 5 "Installation complete!" "Ready to use..."

        return $true
    } catch {
        return $false, $_.Exception.Message
    }
}

# Main wizard logic
function Show-Wizard {
    $script:Wizard = Create-WizardForm
    $currentPage = 0
    $pages = @()

    # Create navigation buttons
    $script:BackButton = New-Object System.Windows.Forms.Button
    $script:BackButton.Text = "< Back"
    $script:BackButton.Location = New-Object System.Drawing.Point(350, 370)
    $script:BackButton.Size = New-Object System.Drawing.Size(80, 30)
    $script:BackButton.Enabled = $false
    $script:Wizard.Controls.Add($script:BackButton)

    $script:NextButton = New-Object System.Windows.Forms.Button
    $script:NextButton.Text = "Next >"
    $script:NextButton.Location = New-Object System.Drawing.Point(440, 370)
    $script:NextButton.Size = New-Object System.Drawing.Size(80, 30)
    $script:NextButton.DialogResult = "OK"
    $script:Wizard.Controls.Add($script:NextButton)

    $script:CancelButton = New-Object System.Windows.Forms.Button
    $script:CancelButton.Text = "Cancel"
    $script:CancelButton.Location = New-Object System.Drawing.Point(520, 370)
    $script:CancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $script:CancelButton.DialogResult = "Cancel"
    $script:Wizard.Controls.Add($script:CancelButton)

    # Page navigation
    $script:NextButton.Add_Click({
        switch ($currentPage) {
            0 { # Welcome -> License
                $currentPage = 1
                $script:BackButton.Enabled = $true
            }
            1 { # License -> Location
                if (-not $script:AgreeCheckBox.Checked) {
                    [System.Windows.Forms.MessageBox]::Show("Please accept the license agreement to continue.", "License Agreement", "OK", "Warning")
                    return
                }
                $currentPage = 2
            }
            2 { # Location -> Install
                if ([string]::IsNullOrWhiteSpace($script:InstallPath)) {
                    [System.Windows.Forms.MessageBox]::Show("Please select an installation location.", "Installation Location", "OK", "Warning")
                    return
                }
                $currentPage = 3
                $script:NextButton.Enabled = $false
                $script:BackButton.Enabled = $false
                $script:CancelButton.Enabled = $false

                # Start installation in background
                $installJob = Start-Job -ScriptBlock ${function:Install-Application}
                $timer = New-Object System.Windows.Forms.Timer
                $timer.Interval = 500
                $timer.Add_Tick({
                    if ($installJob.State -eq "Completed") {
                        $timer.Stop()
                        $result = Receive-Job $installJob
                        Remove-Job $installJob

                        $success = $result[0]
                        $errorMessage = if ($result.Count -gt 1) { $result[1] } else { "" }

                        $currentPage = 4
                        $script:NextButton.Text = "Finish"
                        $script:NextButton.Enabled = $true
                        $script:CancelButton.Text = "Close"
                        $script:CancelButton.Enabled = $true

                        # Show completion page
                        $pages[4] = Show-CompletePage $success $errorMessage
                        $script:Wizard.Controls.Clear()
                        $script:Wizard.Controls.Add($pages[4])
                        Add-NavigationButtons
                    }
                })
                $timer.Start()
            }
            3 { # Install (shouldn't happen)
            }
            4 { # Complete
                if ($script:LaunchCheckBox -and $script:LaunchCheckBox.Checked) {
                    Start-Process "$script:InstallPath\Launch.bat"
                }
                $script:Wizard.Close()
            }
        }

        # Update page display
        if ($currentPage -lt $pages.Count) {
            $script:Wizard.Controls.Clear()
            $script:Wizard.Controls.Add($pages[$currentPage])
            Add-NavigationButtons
        }
    })

    $script:BackButton.Add_Click({
        if ($currentPage -gt 0) {
            $currentPage--
            $script:Wizard.Controls.Clear()
            $script:Wizard.Controls.Add($pages[$currentPage])
            Add-NavigationButtons
        }
    })

    $script:CancelButton.Add_Click({
        if ($currentPage -eq 3) {
            # Installation in progress, confirm cancel
            $result = [System.Windows.Forms.MessageBox]::Show("Installation is in progress. Are you sure you want to cancel?", "Cancel Installation", "YesNo", "Question")
            if ($result -eq "No") { return }
        }
        $script:Wizard.Close()
    })

    function Add-NavigationButtons {
        $script:Wizard.Controls.Add($script:BackButton)
        $script:Wizard.Controls.Add($script:NextButton)
        $script:Wizard.Controls.Add($script:CancelButton)

        # Update button states
        $script:BackButton.Enabled = $currentPage -gt 0 -and $currentPage -lt 3
        $script:NextButton.Enabled = $currentPage -ne 3
        $script:CancelButton.Enabled = $currentPage -ne 4

        switch ($currentPage) {
            0 { $script:NextButton.Text = "Next >" }
            1 { $script:NextButton.Text = "Next >" }
            2 { $script:NextButton.Text = "Install" }
            3 { $script:NextButton.Text = "Installing..." }
            4 { $script:NextButton.Text = "Finish" }
        }
    }

    # Create pages
    $pages += Show-WelcomePage
    $pages += Show-LicensePage
    $pages += Show-LocationPage
    $pages += Show-InstallPage
    $pages += Show-CompletePage

    # Show first page
    $script:Wizard.Controls.Add($pages[0])
    Add-NavigationButtons

    # Check Python first
    $pythonInfo = Find-Python
    if (-not $pythonInfo) {
        [System.Windows.Forms.MessageBox]::Show("Python 3.8 or later is required but not found.`n`nPlease install Python from https://python.org and run this installer again.", "Python Required", "OK", "Error")
        return
    }

    # Show wizard
    $result = $script:Wizard.ShowDialog()

    if ($result -eq "Cancel") {
        Write-Host "Installation cancelled by user."
    }
}

# Run the wizard
Show-Wizard
