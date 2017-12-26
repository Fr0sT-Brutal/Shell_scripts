'Script to download an URL to a file
'  %1 - URL
'  %2 - file
'Works only under cscript for some reason...

'Original source http://www.ericphelps.com/scripting/samples/BinaryDownload/index.htm
'Works only under cscript for some reason...
Function SaveWebBinary(strUrl, strFile) 'As Boolean
    Const adTypeBinary = 1
    Const adSaveCreateOverWrite = 2
    Const ForWriting = 2
    Dim web, varByteArray, strData, strBuffer, lngCounter, ado
    On Error Resume Next
    'Download the file with any available object
    Err.Clear
    Set web = Nothing
    Set web = CreateObject("WinHttp.WinHttpRequest.5.1")
    If web Is Nothing Then Set web = CreateObject("WinHttp.WinHttpRequest")
    If web Is Nothing Then Set web = CreateObject("MSXML2.ServerXMLHTTP")
    If web Is Nothing Then Set web = CreateObject("Microsoft.XMLHTTP")
    web.Open "GET", strURL, False
    web.Send
    If Err.Number <> 0 Then
        WScript.Echo "Err: Get, error " & Err.Number
        SaveWebBinary = False
        Set web = Nothing
        Exit Function
    End If
    If web.Status <> "200" Then
        WScript.Echo "Err: Get, status " & web.Status
        SaveWebBinary = False
        Set web = Nothing
        Exit Function
    End If
    varByteArray = web.ResponseBody
    Set web = Nothing
    'Now save the file with any available method
    On Error Resume Next
    Set ado = Nothing
    Set ado = CreateObject("ADODB.Stream")
    If ado Is Nothing Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set ts = fs.OpenTextFile(strFile, ForWriting, True)
        strData = ""
        strBuffer = ""
        For lngCounter = 0 to UBound(varByteArray)
            ts.Write Chr(255 And Ascb(Midb(varByteArray,lngCounter + 1, 1)))
        Next
        ts.Close
    Else
        ado.Type = adTypeBinary
        ado.Open
        ado.Write varByteArray
        ado.SaveToFile strFile, adSaveCreateOverWrite
        ado.Close
    End If
    SaveWebBinary = True
End Function

' Main

If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Expect 2 arguments: URL and destination file"
    WScript.Quit 1
End if

sUrl = WScript.Arguments.Item(0)
sFile = WScript.Arguments.Item(1)
If Not SaveWebBinary(sUrl, sFile) Then
    WScript.Echo "Error downloading or saving file"
    WScript.Quit 1
End if