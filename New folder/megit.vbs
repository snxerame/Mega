Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strPath
strPath = "C:\Mega"

If Not objFSO.FolderExists(strPath) Then
    objFSO.CreateFolder(strPath)
End If

Dim objXML
Set objXML = CreateObject("MSXML2.XMLHTTP")
objXML.Open "GET", "https://github.com/Foxit9/Mega", False
objXML.Send

If objXML.Status = 200 Then
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.Write objXML.ResponseBody
    objStream.Position = 0
    objStream.SaveToFile strPath & "\downloaded_content.html", 2
    objStream.Close
    Set objStream = Nothing
End If

Set objXML = Nothing
Set objFSO = Nothing
