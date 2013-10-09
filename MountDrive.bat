:: Mount local/remote folder/drive as disk and assign a label to it.
:: %1 - Disk letter to mount to
:: %2 - Path to folder to mount
:: %3 - User
:: %4 - Pass
:: %5 - Label of the mounted disk

@ECHO OFF

SETLOCAL
SET CDir=%~dp0%

:: Remove already mounted disk
NET USE %1 /DELETE 2> NUL
:: Mount
NET USE %1 %2 %4 "/USER:%~3" /PERSISTENT:NO
:: Label the disk if mounted successfully
IF ERRORLEVEL 1 (
	PAUSE
) ELSE (
	CALL cscript.exe "%CDir%\DiskRen.js" %1 %5 > NUL
)