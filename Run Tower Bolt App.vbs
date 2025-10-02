Set objShell = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Path to Anaconda Python (only used to create venv)
pythonExe = "C:\ProgramData\anaconda3\python.exe"

' Resolve this .vbs location
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' App directory where all files are located
appDir = fso.BuildPath(scriptDir, "app")

' Venv directory (inside app folder)
venvDir = fso.BuildPath(appDir, "venv")
venvPython = fso.BuildPath(venvDir, "Scripts\python.exe")

' Python script and requirements paths (inside app folder)
mainScript = fso.BuildPath(appDir, "main.py")
requirementsFile = fso.BuildPath(appDir, "requirements.txt")

' Create PowerShell script for progress window in app folder
psScriptPath = fso.BuildPath(appDir, "launcher_progress.ps1")
Set psFile = fso.CreateTextFile(psScriptPath, True)

psFile.WriteLine "Add-Type -AssemblyName System.Windows.Forms"
psFile.WriteLine "Add-Type -AssemblyName System.Drawing"
psFile.WriteLine ""
psFile.WriteLine "$form = New-Object System.Windows.Forms.Form"
psFile.WriteLine "$form.Text = 'Tower Bolt Application Launcher'"
psFile.WriteLine "$form.Size = New-Object System.Drawing.Size(500, 200)"
psFile.WriteLine "$form.StartPosition = 'CenterScreen'"
psFile.WriteLine "$form.FormBorderStyle = 'FixedDialog'"
psFile.WriteLine "$form.MaximizeBox = $false"
psFile.WriteLine "$form.MinimizeBox = $false"
psFile.WriteLine "$form.TopMost = $true"
psFile.WriteLine ""
psFile.WriteLine "$label = New-Object System.Windows.Forms.Label"
psFile.WriteLine "$label.Location = New-Object System.Drawing.Point(20, 20)"
psFile.WriteLine "$label.Size = New-Object System.Drawing.Size(460, 30)"
psFile.WriteLine "$label.Font = New-Object System.Drawing.Font('Segoe UI', 10)"
psFile.WriteLine "$label.Text = 'Initializing...'"
psFile.WriteLine "$form.Controls.Add($label)"
psFile.WriteLine ""
psFile.WriteLine "$progressBar = New-Object System.Windows.Forms.ProgressBar"
psFile.WriteLine "$progressBar.Location = New-Object System.Drawing.Point(20, 60)"
psFile.WriteLine "$progressBar.Size = New-Object System.Drawing.Size(460, 30)"
psFile.WriteLine "$progressBar.Style = 'Continuous'"
psFile.WriteLine "$form.Controls.Add($progressBar)"
psFile.WriteLine ""
psFile.WriteLine "$detailLabel = New-Object System.Windows.Forms.Label"
psFile.WriteLine "$detailLabel.Location = New-Object System.Drawing.Point(20, 100)"
psFile.WriteLine "$detailLabel.Size = New-Object System.Drawing.Size(460, 50)"
psFile.WriteLine "$detailLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)"
psFile.WriteLine "$detailLabel.Text = ''"
psFile.WriteLine "$form.Controls.Add($detailLabel)"
psFile.WriteLine ""
psFile.WriteLine "$form.Show()"
psFile.WriteLine "$form.Refresh()"
psFile.WriteLine ""
psFile.WriteLine "function Update-Progress {"
psFile.WriteLine "    param($percent, $status, $detail)"
psFile.WriteLine "    $progressBar.Value = $percent"
psFile.WriteLine "    $label.Text = $status"
psFile.WriteLine "    $detailLabel.Text = $detail"
psFile.WriteLine "    $form.Refresh()"
psFile.WriteLine "    [System.Windows.Forms.Application]::DoEvents()"
psFile.WriteLine "}"
psFile.WriteLine ""
psFile.WriteLine "$venvExists = Test-Path """ & venvDir & """"
psFile.WriteLine "$needsVenv = -not $venvExists"
psFile.WriteLine "$needsRequirements = Test-Path """ & requirementsFile & """"
psFile.WriteLine ""
psFile.WriteLine "$totalSteps = 1"
psFile.WriteLine "if ($needsVenv) { $totalSteps++ }"
psFile.WriteLine "if ($needsRequirements) { $totalSteps++ }"
psFile.WriteLine "$currentStep = 0"
psFile.WriteLine ""
psFile.WriteLine "try {"
psFile.WriteLine "    if ($needsVenv) {"
psFile.WriteLine "        $currentStep++"
psFile.WriteLine "        Update-Progress -percent (($currentStep / $totalSteps) * 100) -status 'Creating virtual environment...' -detail 'This may take a moment...'"
psFile.WriteLine "        $result = & """ & pythonExe & """ -m venv """ & venvDir & """ 2>&1"
psFile.WriteLine "        if ($LASTEXITCODE -ne 0) { throw 'Failed to create virtual environment' }"
psFile.WriteLine "    }"
psFile.WriteLine ""
psFile.WriteLine "    if ($needsRequirements) {"
psFile.WriteLine "        $currentStep++"
psFile.WriteLine "        Update-Progress -percent (($currentStep / $totalSteps) * 100) -status 'Installing requirements...' -detail 'Installing packages: pandas, PyQt5, matplotlib, numpy, openpyxl'"
psFile.WriteLine "        $result = & """ & venvPython & """ -m pip install -q -r """ & requirementsFile & """ 2>&1"
psFile.WriteLine "        if ($LASTEXITCODE -ne 0) { throw 'Failed to install requirements' }"
psFile.WriteLine "    }"
psFile.WriteLine ""
psFile.WriteLine "    $currentStep++"
psFile.WriteLine "    Update-Progress -percent 100 -status 'Launching application...' -detail 'Starting Tower Bolt Application'"
psFile.WriteLine "    Start-Sleep -Milliseconds 500"
psFile.WriteLine "    $form.Close()"
psFile.WriteLine ""
psFile.WriteLine "    & """ & venvPython & """ """ & mainScript & """"
psFile.WriteLine "} catch {"
psFile.WriteLine "    [System.Windows.Forms.MessageBox]::Show(""Error: $($_.Exception.Message)"", 'Launch Error', 'OK', 'Error')"
psFile.WriteLine "    $form.Close()"
psFile.WriteLine "}"

psFile.Close

' Run PowerShell script
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psScriptPath & """", 0, True

' Clean up PowerShell script
If fso.FileExists(psScriptPath) Then
    fso.DeleteFile psScriptPath
End If