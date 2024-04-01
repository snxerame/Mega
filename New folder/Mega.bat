@echo off

rem Set the URL of the GitHub repository's zip file
set "URL=https://github.com/Foxit9/Mega/archive/main.zip"

rem Set the download path and zip file name
set "DOWNLOAD_PATH=C:\Mega"
set "ZIP_FILE=%DOWNLOAD_PATH%\Mega.zip"

rem Create the download path if it doesn't exist
if not exist "%DOWNLOAD_PATH%" mkdir "%DOWNLOAD_PATH%"

rem Download the zip file using PowerShell
powershell -Command "(New-Object Net.WebClient).DownloadFile('%URL%', '%ZIP_FILE%')"

rem Check if the download was successful
if exist "%ZIP_FILE%" (
    rem Extract the contents of the zip file using PowerShell
    powershell -Command "Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%DOWNLOAD_PATH%'"

    rem Remove the zip file after extraction
    del "%ZIP_FILE%"

    echo Repository downloaded successfully to: "%DOWNLOAD_PATH%\Mega-main"
) else (
    echo Failed to download the repository.
)
