# ========================================
# Tower Bolt Tension Data Tool - Standalone GUI Installer
# ========================================
# Downloads from GitHub, creates venv with system Python, installs dependencies
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Configuration
$GitHubUser = "benchaney53"
$RepoName = "VestasTensionDataTool"
$Branch = "master"
$AppName = "Tower Bolt Tension Data Tool"

# Hardcoded system Python path (same as VBS launcher)
$SystemPython = "C:\ProgramData\anaconda3\python.exe"

# Global variables
$script:InstallPath = ""
$script:CurrentStep = 0
$script:TotalSteps = 6
$script:Wizard = $null
$script:ProgressBar = $null
$script:StatusLabel = $null
$script:InstallResult = $null
$script:Timer = $null

# Create the installer form
function Create-InstallerForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$AppName - Installation"
    $form.Size = New-Object System.Drawing.Size(500, 300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Icon = [System.Drawing.SystemIcons]::Application
    $form.BackColor = [System.Drawing.Color]::White
    $form.TopMost = $true

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Installing $AppName"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 40)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 114, 198)  # Vestas blue
    $form.Controls.Add($titleLabel)

    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Preparing installation..."
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $statusLabel.Location = New-Object System.Drawing.Point(20, 70)
    $statusLabel.Size = New-Object System.Drawing.Size(460, 60)
    $form.Controls.Add($statusLabel)
    $script:StatusLabel = $statusLabel

    # Progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 140)
    $progressBar.Size = New-Object System.Drawing.Size(460, 25)
    $progressBar.Style = "Continuous"
    $progressBar.Minimum = 0
    $progressBar.Maximum = 100
    $form.Controls.Add($progressBar)
    $script:ProgressBar = $progressBar

    # Detail label
    $detailLabel = New-Object System.Windows.Forms.Label
    $detailLabel.Text = ""
    $detailLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $detailLabel.ForeColor = [System.Drawing.Color]::Gray
    $detailLabel.Location = New-Object System.Drawing.Point(20, 175)
    $detailLabel.Size = New-Object System.Drawing.Size(460, 40)
    $form.Controls.Add($detailLabel)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(390, 230)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.DialogResult = "Cancel"
    $form.Controls.Add($cancelButton)

    $form.CancelButton = $cancelButton

    return $form
}

# Update progress in GUI
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

# Perform the actual installation (runs in main thread)
function Install-Application {
    try {
        Update-Progress 1 "Checking system Python..." "Verifying Python installation..."

        # Check if system Python exists
        if (-not (Test-Path $SystemPython)) {
            throw "System Python not found at: $SystemPython`nPlease ensure Anaconda is installed."
        }

        # Verify Python works
        $pythonVersion = & $SystemPython "--version" 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "System Python is not working. Please check your Anaconda installation."
        }
        Update-Progress 1 "Found Python: $($pythonVersion.Trim())" "Using system Python for venv creation"

        # Ask for installation location
        Update-Progress 2 "Selecting installation location..." "Choosing where to install $AppName"

        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select installation folder for $AppName"
        $folderBrowser.SelectedPath = "$env:USERPROFILE\Desktop"
        $folderBrowser.ShowNewFolderButton = $true

        if ($folderBrowser.ShowDialog() -eq "OK") {
            $script:InstallPath = $folderBrowser.SelectedPath
        } else {
            throw "Installation cancelled by user."
        }

        Update-Progress 2 "Installation location: $script:InstallPath" "Selected folder for installation"

        # Create installation directory
        Update-Progress 3 "Creating installation directory..." $script:InstallPath
        if (Test-Path $script:InstallPath) {
            Remove-Item $script:InstallPath -Recurse -Force
        }
        New-Item -ItemType Directory -Path $script:InstallPath -Force | Out-Null

        # Download from GitHub
        Update-Progress 4 "Downloading from GitHub..." "Fetching latest version..."
        $zipUrl = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
        $tempZip = "$env:TEMP\$RepoName.zip"
        $tempExtract = "$env:TEMP\$RepoName-extract"

        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }

        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
        } catch {
            throw "Failed to download from GitHub. Please check your internet connection."
        }

        # Extract
        Update-Progress 4 "Extracting files..." "Unpacking application files..."
        Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
        $extractedFolder = Get-ChildItem $tempExtract | Select-Object -First 1
        Copy-Item "$($extractedFolder.FullName)\*" -Destination $script:InstallPath -Recurse -Force

        # Cleanup
        Remove-Item $tempZip -Force
        Remove-Item $tempExtract -Recurse -Force

        # Create virtual environment using system Python
        Update-Progress 5 "Setting up Python environment..." "Creating virtual environment..."
        $venvPath = "$script:InstallPath\app\venv"

        & $SystemPython -m venv $venvPath
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to create virtual environment. Please check your Python installation."
        }

        # Install dependencies using venv
        Update-Progress 5 "Installing dependencies..." "Installing required packages..."
        $requirementsPath = "$script:InstallPath\app\requirements.txt"
        if (Test-Path $requirementsPath) {
            & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
            & "$venvPath\Scripts\pip.exe" install -r $requirementsPath --quiet
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to install dependencies. Please check your internet connection."
            }
        }

        # Create launcher
        Update-Progress 6 "Creating launcher..." "Setting up application launcher..."
        $launcherContent = @"
@echo off
REM $AppName Launcher
cd /d "$script:InstallPath\app"
"$venvPath\Scripts\python.exe" "$script:InstallPath\app\main.py"
pause
"@
        $launcherContent | Out-File -FilePath "$script:InstallPath\Launch.bat" -Encoding UTF8

        # Create installation info
        $infoContent = @"
$AppName - Installation Info
================================
Installation Date: $(Get-Date)
System Python Used: $SystemPython
Python Version: $($pythonVersion.Trim())
Installation Path: $script:InstallPath
Virtual Environment: $venvPath

This file is for troubleshooting purposes only.
"@
        $infoContent | Out-File -FilePath "$script:InstallPath\installation_info.txt" -Encoding UTF8

        # Create desktop shortcut
        Update-Progress 6 "Creating desktop shortcut..." "Adding to desktop..."
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$AppName.lnk")
        $Shortcut.TargetPath = "$script:InstallPath\Launch.bat"
        $Shortcut.WorkingDirectory = "$script:InstallPath\app"
        $Shortcut.Description = $AppName
        $Shortcut.Save()

        Update-Progress 6 "Installation complete!" "Ready to use..."

        return $true
    } catch {
        return $false, $_.Exception.Message
    }
}

# Main installation logic
function Start-Installation {
    $script:Wizard = Create-InstallerForm

    # Handle form closing
    $script:Wizard.Add_FormClosing({
        if ($script:Timer -and $script:Timer.Enabled) {
            $result = [System.Windows.Forms.MessageBox]::Show("Installation is in progress. Are you sure you want to cancel?", "Cancel Installation", "YesNo", "Question")
            if ($result -eq "No") {
                $_.Cancel = $true
            }
        }
    })

    # Timer to check installation progress
    $script:Timer = New-Object System.Windows.Forms.Timer
    $script:Timer.Interval = 100

    $script:Timer.Add_Tick({
        if ($script:InstallResult -ne $null) {
            $script:Timer.Stop()
            $script:Timer.Dispose()

            if ($script:InstallResult -eq $true) {
                [System.Windows.Forms.MessageBox]::Show("$AppName has been successfully installed!`n`nInstallation location: $script:InstallPath`n`nYou can now launch the application from the desktop shortcut.", "Installation Complete", "OK", "Information")
            } else {
                $errorMessage = $script:InstallResult[1]
                [System.Windows.Forms.MessageBox]::Show("Installation failed:`n`n$errorMessage", "Installation Error", "OK", "Error")
            }

            $script:Wizard.Close()
        }
    })

    # Show the form (non-modal)
    $script:Wizard.Show()

    # Start installation immediately
    $script:Timer.Start()

    try {
        # Run installation in background
        $installJob = Start-Job -ScriptBlock {
            try {
                $SystemPython = "C:\ProgramData\anaconda3\python.exe"
                $GitHubUser = "benchaney53"
                $RepoName = "VestasTensionDataTool"
                $Branch = "master"
                $AppName = "Tower Bolt Tension Data Tool"

                # Check if system Python exists
                if (-not (Test-Path $SystemPython)) {
                    throw "System Python not found at: $SystemPython"
                }

                # Verify Python works
                $pythonVersion = & $SystemPython "--version" 2>$null
                if ($LASTEXITCODE -ne 0) {
                    throw "System Python is not working"
                }

                # Download from GitHub
                $zipUrl = "https://github.com/$GitHubUser/$RepoName/archive/refs/heads/$Branch.zip"
                $tempZip = "$env:TEMP\$RepoName.zip"
                $tempExtract = "$env:TEMP\$RepoName-extract"

                if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
                if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }

                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing

                # Extract
                Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
                $extractedFolder = Get-ChildItem $tempExtract | Select-Object -First 1

                # Copy to a temporary location first, then move to final location
                $tempInstallPath = "$env:TEMP\TowerBoltTemp"
                if (Test-Path $tempInstallPath) { Remove-Item $tempInstallPath -Recurse -Force }
                Copy-Item "$($extractedFolder.FullName)\*" -Destination $tempInstallPath -Recurse -Force

                # Cleanup
                Remove-Item $tempZip -Force
                Remove-Item $tempExtract -Recurse -Force

                # Create virtual environment using system Python
                $venvPath = "$tempInstallPath\app\venv"
                & $SystemPython -m venv $venvPath
                if ($LASTEXITCODE -ne 0) {
                    throw "Failed to create virtual environment"
                }

                # Install dependencies using venv
                $requirementsPath = "$tempInstallPath\app\requirements.txt"
                if (Test-Path $requirementsPath) {
                    & "$venvPath\Scripts\python.exe" -m pip install --upgrade pip --quiet
                    & "$venvPath\Scripts\pip.exe" install -r $requirementsPath --quiet
                    if ($LASTEXITCODE -ne 0) {
                        throw "Failed to install dependencies"
                    }
                }

                return $true, $tempInstallPath
            } catch {
                return $false, $_.Exception.Message
            }
        }

        # Wait for job to complete and get result
        Wait-Job $installJob | Out-Null
        $jobResult = Receive-Job $installJob
        Remove-Job $installJob

        $success = $jobResult[0]
        $data = $jobResult[1]

        if ($success) {
            $tempPath = $data
            # Move from temp location to final location (use desktop as default)
            $script:InstallPath = "$env:USERPROFILE\Desktop\$AppName"
            if (Test-Path $script:InstallPath) {
                Remove-Item $script:InstallPath -Recurse -Force
            }
            Move-Item $tempPath $script:InstallPath -Force

            # Create launcher files
            $venvPath = "$script:InstallPath\app\venv"
            $launcherContent = @"
@echo off
REM $AppName Launcher
cd /d "$script:InstallPath\app"
"$venvPath\Scripts\python.exe" "$script:InstallPath\app\main.py"
pause
"@
            $launcherContent | Out-File -FilePath "$script:InstallPath\Launch.bat" -Encoding UTF8

            # Create desktop shortcut
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\$AppName.lnk")
            $Shortcut.TargetPath = "$script:InstallPath\Launch.bat"
            $Shortcut.WorkingDirectory = "$script:InstallPath\app"
            $Shortcut.Description = $AppName
            $Shortcut.Save()

            $script:InstallResult = $true
        } else {
            $script:InstallResult = $false, $data
        }

    } catch {
        $script:InstallResult = $false, $_.Exception.Message
    }

    # Keep the application running to process timer events
    [System.Windows.Forms.Application]::Run($script:Wizard)
}

# Run the installer
Start-Installation
