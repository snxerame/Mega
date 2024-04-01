Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Specify the path to the MegaCMD installer executable
installerPath = "C:\Users\shree\OneDrive\Desktop\MEGAcmdSetup64.exe"

' Specify the directory where you want MegaCMD to be installed
installDirectory = "C:\MegaCMD"

' Build the command to run the installer silently
cmd = installerPath & " /S /D=" & installDirectory

' Run the installer silently
objShell.Run cmd, 0, True

' Check if MegaCMD installation directory exists
If objFSO.FolderExists(installDirectory) Then
    WScript.Echo "MegaCMD installed successfully."
Else
    WScript.Echo "Failed to install MegaCMD."
End If
