' Wrapper script to run an application with elevated rights (Vista+)
' You'll be promted for permission
' Run it with wscript.exe:
'    wscript.exe <path>\sudo.vbs <app to run> <parameters for the app>
' If you want to pass parameters with spaces to an app, you'll have to use apostrophes
'   instead of quotes as the quotes get eaten regardless the quantity (usually doubling
'   should work but not in the case of Windows Scripting Host). Ex.:
'     wscript.exe sudo.vbs cmd /k dir 'c:\program files'
' Unfortunately this means you'll won't be able to use strings/paths/whatever with apostrophes
'   but it's a rather rare case.

If WScript.Arguments.Count = 0 then
    WScript.Quit
End if

Set shApp = CreateObject("Shell.Application")
' Re-construct argumants string without 0th item and with properly placed quotes
Args = vbNullString
For i = 1 to WScript.Arguments.Count - 1
    Arg = Replace(WScript.Arguments.Item(i), "'", """")
    Args = Args & Arg & " "
Next

'Uncomment this line if you'll need to debug
'MsgBox Args

shApp.ShellExecute WScript.Arguments.Item(0), Args, "", "runas", 1