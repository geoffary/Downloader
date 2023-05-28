Dim destination1, destination2

Set objShell = CreateObject("WScript.Shell")
startupFolder = objShell.SpecialFolders("Startup")
destination1 = startupFolder & "\WebHelperService.exe"

appDataFolder = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
folderPath = appDataFolder & "\WindowsBackgroundProcess"
destination2 = folderPath & "\WebServiceHelper.dll"

' Create a reference to the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if the files exist and delete them
If objFSO.FileExists(destination1) Then
    objFSO.DeleteFile(destination1)
End If

If objFSO.FileExists(destination2) Then
    objFSO.DeleteFile(destination2)
End If

' Delete the folder if it's empty
If objFSO.FolderExists(folderPath) Then
    If objFSO.GetFolder(folderPath).Files.Count = 0 Then
        objFSO.DeleteFolder(folderPath)
    End If
End If

strKeyPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\"
strValueName = "MyProgram"
Set objWshReg= CreateObject("WScript.Shell")

' Check if the registry value exists and delete it
On Error Resume Next
objWshReg.RegRead strKeyPath & strValueName

If Err.Number = 0 Then
    objWshReg.RegDelete strKeyPath & strValueName
End If
