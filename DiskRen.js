sUsage =
"Script changes label of a specified disk\n" +
"  arg1 - disk letter\n" +
"  arg2 - new label\n" +
"Ex.: cscript.exe DiskRen.js X: \"New Name\"";

var Args = WScript.Arguments;
// no args - print usage
if (Args.length != 2)
{
	WScript.Echo(sUsage);
	WScript.Quit(1);
}
var oShell = new ActiveXObject("Shell.Application");
oShell.NameSpace(Args.item(0)).Self.Name = Args.item(1);