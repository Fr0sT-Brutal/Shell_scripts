:: Download an URL to a file
::   %1 - URL
::   %2 - file

@ECHO OFF

SETLOCAL

IF .%1.==.. GOTO :NoParam
IF .%2.==.. GOTO :NoParam

CALL powershell -Command "(New-Object Net.WebClient).DownloadFile('%~1', '%~2')"

:: PS 3.0: powershell -Command "Invoke-WebRequest %1 -OutFile %2"

GOTO :EOF

:NoParam
ECHO Usage: GetURLps %URL% %File%