On Error Resume Next

Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' URL of the MegaCMD installation package
strURL = "https://mega.nz/cmd"
strDownloadPath = "C:\MegaC"

' Create the MegaC directory if it doesn't exist
If Not objFSO.FolderExists(strDownloadPath) Then
    objFSO.CreateFolder(strDownloadPath)
End If

' Download the MegaCMD installer to the MegaC directory
objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.Send

If objXMLHTTP.Status = 200 Then
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objXMLHTTP.ResponseBody
    objStream.Position = 0

    ' Save the downloaded installer to the MegaC directory
    strInstallerPath = strDownloadPath & "\MegaCMD.exe"
    objStream.SaveToFile strInstallerPath, 2
    objStream.Close

    ' Run the installer to extract files without installing
    objShell.NameSpace(strDownloadPath).CopyHere objShell.NameSpace(strInstallerPath).Items

    ' Wait for the extraction process to complete
    WScript.Sleep 5000

    ' Clean up the installer file
    objFSO.DeleteFile strInstallerPath
End If

If Err.Number <> 0 Then
    WScript.Echo "An error occurred: " & Err.Description
End If