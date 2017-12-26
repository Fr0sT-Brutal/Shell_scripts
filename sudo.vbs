' Wrapper script to run an application with elevated rights (Vista+)
' You'll be promted for permission
' Run it with wscript.exe:
'    wscript.exe <path>\sudo.vbs <app to run> <parameters for the app>
' NOTE for parameters with spaces. 
'   Windows Scripting Host eats all the quote chars (") you pass inside parameters so you
'   couldn't use quotes inside quotes. Use single quote chars (') instead. Ex.:
'   wscript.exe sudo.vbs "d:\app.exe" "/batch '/path=c:\program files'"

' !!! Remember that caller's current dir WILL NOT be transferred to a launched app. You MUST use
' full paths to a launched app/script (unless its path is in the %PATH%). Also remember that
' a launched app will have %Windows%\System32 as current dir.

If WScript.Arguments.Count = 0 Then
    WScript.Quit 1
End if

Const ch_Quote = """"
Const ch_SglQuote = "'"

Set shApp = CreateObject("Shell.Application")
' Re-construct arguments string without 0th item and with properly placed quotes
Args = vbNullString
For i = 1 to WScript.Arguments.Count - 1
    Arg = WScript.Arguments.Item(i)
    ' If argument contains spaces, quote it
    If InStr(Arg, " ") > 0 Then
    	Arg = ch_Quote & Arg & ch_Quote
    End If
    ' Replace ' with "" (use with care!)
    Arg = Replace(Arg, ch_SglQuote, ch_Quote & ch_Quote)
    Args = Args & Arg & " "
Next
App = WScript.Arguments.Item(0)

'Uncomment this line if you'll need to debug
'MsgBox("App: " & App & vbNewLine & "Args: " & Args)

shApp.ShellExecute App, Args, "", "runas", 1