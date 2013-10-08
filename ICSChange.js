sUsage =
"Script for changing Internet connection sharing (ICS).\n" +
"NB: you MUST configure a network to share connection with before using the script.\n" +
"Usage (Run as admin under Vista and greater!):\n" +
"  toggle_ics.js [list | <Adapter index> [ON|OFF] | <Adapter name> [ON|OFF]]\n" +
"  * No parameters           get this text\n" +
"  * list                    shows a list of available adapters as \n" +
"                            '<index>. <name> @ <device>'\n" +
"  * <Adapter index>         toggles ICS on an adapter with specified index\n" +
"                            (may vary depending on current config)\n" +
"  * <Adapter index> ON|OFF  ensures specified ICS state on an adapter with\n" +
"                            specified index\n" +
"  * <Adapter name>          toggles ICS on an adapter with specified name\n" +
"  * <Adapter name> ON|OFF   ensures specified ICS state on an adapter with\n" +
"                            specified name\n" +
"!! For the latter two, at first check the list of available adapters to get\n" +
"proper value\n";

ICSSHARINGTYPE_PUBLIC = 0;

var objShare = new ActiveXObject("HNetCfg.HNetShare.1"); // share manager (INetSharingManager intf)
var conns = new Array(); // all available network adapters
var connection = null; // current adapter to work with
var desiredState = null; // desired ICS state

// get all available adapters and push them into an array
function EnumConns()
{
	// get all connections (INetSharingEveryConnectionCollection intf)
	var objEveryColl = objShare.EnumEveryConnection;
	if (objEveryColl == null || objEveryColl.Count == 0) return;

	// convert to built-in Java-style enumerator
	var e = new Enumerator(objEveryColl);
	e.moveFirst();
	for (; !e.atEnd(); e.moveNext())
	{
		// get an INetConnection interface (not an ole-automation object)
		var objNetConn = e.item();
		if (objNetConn != null)	   
			conns.push(objNetConn);	   
	}
}

// show the list of found adapters
function ShowConns()
{
	var str = "";
	for (var i = 0; i < conns.length; i++)
	{
		// find the right connection, by examining the NetConnectionProps
		var objNCProps = objShare.NetConnectionProps(conns[i]);
		str += (i + ". " + objNCProps.Name + " @ " + objNCProps.DeviceName + "\n");
	}
	WScript.Echo(str);
}

// determine args and perform required actions
function checkArgs()
{
	objArgs = WScript.Arguments;

	// no args - print usage
	if (objArgs.length == 0)
	{
		WScript.Echo(sUsage);
		WScript.Quit(1);
	}
	
	// arg #1 is "list" - show adapters
	if (objArgs(0) == "list")
	{
		EnumConns();
		ShowConns();
		WScript.Quit(0);
	}
	
	// otherwise consider arg #1 is adapter index/name; find the object
	EnumConns();
	var idx = parseInt(objArgs(0));
	// arg = index
	if (!isNaN(idx))
		connection = conns[idx];
	// arg = name
	else
	{
		for (var i = 0; i < conns.length; i++)
		{
			// find the right connection, by examining the NetConnectionProps
			var objNCProps = objShare.NetConnectionProps(conns[i]);
			if (objNCProps.Name == objArgs(0))
			{
				connection = conns[i];
				break;
			}
		}	
	}
	// check
	if (connection == null)
	{
		WScript.Echo("Cannot find specified adapter '"+objArgs(0)+"'");
		WScript.Quit(1);
	}
	
	// there's arg #2, check if it is a desired ICS state
	if (objArgs.length >= 2)
	{
		if (objArgs(1).toLowerCase() == "on")
			desiredState = true;
		else if (objArgs(1).toLowerCase() == "off")
			desiredState = false;
		else
		{
			WScript.Echo("Invalid ICS state specified: '"+objArgs(1)+"', use ON or OFF");
			WScript.Quit(1);
		}
	}
}

// Change ICS state for current adapter.
//   aState == null: toggle ICS state ( state_New = not state_Old )
//   aState == true/false: ensure ICS equals to aState. If it is, no change is performed
function setICSState(aState)
{
	var objShareCfg = objShare.INetSharingConfigurationForINetConnection(connection);
	if (aState == null || objShareCfg.SharingEnabled != aState)
	{	
		if (objShareCfg.SharingEnabled)
			objShareCfg.DisableSharing();
		else
			objShareCfg.EnableSharing(ICSSHARINGTYPE_PUBLIC);
	}
	return objShareCfg.SharingEnabled;
}

// *** Main starts here ***
checkArgs();
var res = setICSState(desiredState);
var objNCProps = objShare.NetConnectionProps(connection);
WScript.Echo("ICS for adapter '" + objNCProps.Name + "' is now " + (res ? "ENABLED" : "DISABLED" ));