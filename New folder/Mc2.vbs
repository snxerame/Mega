On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Path to the MegaCMD installer
strInstallerPath = "C:\MegaC\MEGAcmdSetup64.exe"

' Ensure the installer file exists
If objFSO.FileExists(strInstallerPath) Then
    ' Run the installer to extract files without installing
    objShell.NameSpace("C:\MegaC").CopyHere objShell.NameSpace(strInstallerPath).Items

    ' Wait for the extraction process to complete
    WScript.Sleep 5000

    ' Clean up the installer file (optional)
    ' objFSO.DeleteFile strInstallerPath
Else
    WScript.Echo "Installer file not found."
End If

If Err.Number <> 0 Then
    WScript.Echo "An error occurred: " & Err.Description
End If
