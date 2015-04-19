' Settings
strDownloadDir = "target"
cmderURL = "https://github.com/bliker/cmder/releases/download/v1.1.4.1/cmder_mini.zip"

' Globals
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Setup download directory
If objFSO.FolderExists(strDownloadDir) Then
  intAnswer = Msgbox("Download directory '" & strDownloadDir & "' already exists; do you want to delete it?", vbYesNo, "Delete '" & strDownloadDir & "'?")
  If intAnswer = vbYes Then
    objFSO.DeleteFolder strDownloadDir
  Else
    Set objFSO = Nothing
    WScript.Quit
  End If
End If
objFSO.CreateFolder strDownloadDir

' Download function
Function downloadFile(strDescription, strURL, strFileName)
  strFilePath=strDownloadDir & "\" & strFileName
  Wscript.Echo "Downloading " & strDescription & "..."

  Set objXMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
  objXMLHTTP.open "GET", strURL, false
  objXMLHTTP.send()

  If objXMLHTTP.Status = 200 Then
    Set objADOStream = CreateObject("ADODB.Stream")
    objADOStream.Open
    objADOStream.Type = 1 'adTypeBinary

    objADOStream.Write objXMLHTTP.ResponseBody
    objADOStream.Position = 0    'Set the stream position to the start
    objADOStream.SaveToFile strFilePath
    objADOStream.Close
    Set objADOStream = Nothing
  End if

  Set objXMLHTTP = Nothing
  Wscript.Echo "... Saved to " & strFilePath
  downloadFile = strDescription
End Function

' Perform downloads
downloadFile "Cmder console", cmderURL, "cmder_mini.zip"

' Cleanup
Set objFSO = Nothing
Wscript.Echo "File downloads complete."
