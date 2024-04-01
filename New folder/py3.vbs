' Define variables
Dim objShell
Dim objFSO
Dim strPythonInstallerURL
Dim strDownloadFolder
Dim strPythonInstallerPath
Dim strExtractedFolder
Dim strPythonInstallDir

' Set Python installer URL
strPythonInstallerURL = "https://www.python.org/ftp/python/3.10.0/python-3.10.0-amd64.exe"

' Set download folder for Python installer
strDownloadFolder = "C:\PythonDownload"

' Set download path for Python installer
strPythonInstallerPath = strDownloadFolder & "\python-installer.exe"

' Set Python installation directory
strPythonInstallDir = "C:\Python"

' Create objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create download folder if it doesn't exist
If Not objFSO.FolderExists(strDownloadFolder) Then
    objFSO.CreateFolder(strDownloadFolder)
End If

' Download Python installer
objShell.Run "powershell -command ""(New-Object Net.WebClient).DownloadFile('" & strPythonInstallerURL & "', '" & strPythonInstallerPath & "')""", 0, True

' Check if download was successful
If objFSO.FileExists(strPythonInstallerPath) Then
    ' Run Python installer silently
    objShell.Run chr(34) & strPythonInstallerPath & chr(34) & " /quiet InstallAllUsers=0 TargetDir=" & chr(34) & strPythonInstallDir & chr(34), 0, True

    ' Display message box to indicate installation completed
    MsgBox "Python installation completed successfully.", vbInformation, "Installation Complete"
Else
    MsgBox "Failed to download Python installer.", vbExclamation, "Download Failed"
End If

' Clean up
Set objShell = Nothing
Set objFSO = Nothing
