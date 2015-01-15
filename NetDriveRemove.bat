:: Interactive script that lists all LAN connection (net drives) and allows to delete them.
:: Drives could be mounted or not. Not mounted drives could cause confusing as they're not 
:: visible in "My computer > Unmount net drive" and make "multiple connection" error appear.

@ECHO OFF

SETLOCAL
SETLOCAL ENABLEDELAYEDEXPANSION

:OutputList

:: List connections only (no excess words) with numbering

set Cnt=0

echo Current LAN connections:
echo.

for /f "delims=" %%s in ('net use ^| findstr "\\"') do (
	set /a Cnt=!Cnt!+1
	echo !Cnt!.   %%s
)

:: Exit if no connections
if !Cnt!==0 (
	echo No LAN connections
	pause
	goto :EOF
)
echo.

:: Request connection number

:: (!) "set /p" won't assign empty value to a variable if nothing is entered
set Num=
set /p Num=Input connection number to delete, press ENTER to exit: 

if .%Num%.==.. goto :EOF

:: Delete connection that goes under the given number
:: (!) "net use" could output
::   OK    V:    \\serv_name\...    => delete command will be "net use V: /delete"
::   OK          \\serv_name\...    => delete command will be "net use \\serv_name\... /delete"
:: so we always take 2nd token

set Cnt=0

for /f "tokens=2" %%s in ('net use ^| findstr "\\"') do (
	set /a Cnt=!Cnt!+1
	if !Cnt!==!Num! (
		net use %%s /delete
	)
)

goto :OutputList