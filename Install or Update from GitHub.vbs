' ========================================
' GitHub Repository Install/Update Script
' NO GIT REQUIRED - Downloads ZIP from GitHub
' ========================================

' Configuration - EDIT THESE VALUES
username = "benchaney53"           ' GitHub username
repository = "VestasTensionDataTool"       ' Repository name
branch = "master"              ' Branch name (usually "main" or "master")
repoFolderName = "repository"   ' Local folder name for the repo

' ========================================
' Script Logic - Don't edit below unless needed
' ========================================

' Construct the ZIP download URL
zipURL = "https://github.com/" & username & "/" & repository & "/archive/refs/heads/" & branch & ".zip"

' Get the directory where this script is located
Set objFSO = CreateObject("Scripting.FileSystemObject")
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
repoPath = objFSO.BuildPath(scriptPath, repoFolderName)
zipPath = objFSO.BuildPath(scriptPath, "temp_download.zip")
tempExtractPath = objFSO.BuildPath(scriptPath, "temp_extract")

' Check if repository folder already exists
If objFSO.FolderExists(repoPath) Then
    response = MsgBox("Repository found. Do you want to update it?" & vbCrLf & vbCrLf & _
                      "This will replace all files with the latest version from GitHub.", _
                      vbYesNo + vbQuestion, "Update Repository")
    If response = vbNo Then
        WScript.Echo "Update cancelled."
        WScript.Quit
    End If
    WScript.Echo "Updating repository..."
Else
    WScript.Echo "Installing repository..."
End If

' Download the ZIP file
WScript.Echo "Downloading from GitHub..."
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
objHTTP.Open "GET", zipURL, False
objHTTP.Send

If objHTTP.Status = 200 Then
    ' Save the downloaded ZIP
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 1 ' Binary
    objStream.Open
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile zipPath, 2 ' Overwrite
    objStream.Close
    Set objStream = Nothing
    
    WScript.Echo "Download complete. Extracting files..."
    
    ' Delete old repo folder if it exists
    If objFSO.FolderExists(repoPath) Then
        objFSO.DeleteFolder repoPath, True
    End If
    
    ' Create temp extract folder
    If objFSO.FolderExists(tempExtractPath) Then
        objFSO.DeleteFolder tempExtractPath, True
    End If
    objFSO.CreateFolder tempExtractPath
    
    ' Extract ZIP file
    Set objShell = CreateObject("Shell.Application")
    Set zipFile = objShell.NameSpace(zipPath)
    Set extractFolder = objShell.NameSpace(tempExtractPath)
    
    extractFolder.CopyHere zipFile.Items, 4 + 16 ' 4=no progress dialog, 16=yes to all
    
    ' Wait for extraction to complete
    WScript.Sleep 2000
    
    ' Find the extracted folder (GitHub adds "-branchname" to folder)
    Set tempFolder = objFSO.GetFolder(tempExtractPath)
    Set subFolders = tempFolder.SubFolders
    
    For Each subFolder In subFolders
        ' Move the extracted folder to the target location
        objFSO.MoveFolder subFolder.Path, repoPath
        Exit For
    Next
    
    ' Clean up temporary files
    objFSO.DeleteFile zipPath
    objFSO.DeleteFolder tempExtractPath, True
    
    WScript.Echo "Success! Repository saved to:" & vbCrLf & repoPath
    
Else
    WScript.Echo "Error downloading repository. HTTP Status: " & objHTTP.Status & vbCrLf & _
                 "Please check:" & vbCrLf & _
                 "- Username, repository, and branch names are correct" & vbCrLf & _
                 "- Repository is public (or you have access)" & vbCrLf & _
                 "- You have an internet connection"
End If

' Clean up
Set objHTTP = Nothing
Set objShell = Nothing
Set objFSO = Nothing

WScript.Echo vbCrLf & "Press OK to exit."

