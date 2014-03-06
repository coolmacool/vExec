vExec
=====

proof of concept

Initial work on a C# WMI port for interactive remote command execution

Allows running of net use command with network resources available to logged on user 


####Example
--------------------


```c#
// Instantiation
vExec ve = new vExec("MACHINEADDRESS", "username", "password");

// Subscribe to OutputUpdated Event
ve.OutputUpdated += new EventHandler(ve_OutputUpdated);

private void ve_OutputUpdated(object sender, EventArgs e)
{
	// Update UI from different thread
	Invoke((Action)delegate { txtOutput.Text = ve.CmdOutput; });
}

private void button1_Click(object sender, EventArgs e)
{      
	// Run Method to begin command execution excepts command as string
	// Excample command: ipconfig
	
	ve.Run("ipconfig");
	
	//OR
	// Call the sub methods individually 
	
	/* 
	ve.RemoteCommand = "ipconfig";
        ve.ConnectWMI();
        ve.LaunchRemoteProcess();    
        */
}
```
####For WPF Applications

Replace Invoke with Dispatcher.Invoke

```c#
Dispatcher.Invoke((Action)delegate { txtOutput.Text = ve.CmdOutput; });
```

==================
####Properties
```c#
string RemoteComputer
string RemoteCommand
string UserName
string UserPass
string CmdOutput
```


####Public Methods
```c#
void Run(string command)	// Calls ConnectWMI() & LaunchRemoteProcess()
```
Methods for use when wanting to seperate connection and command execution calls
```c#
// Makes connection to remote machine 
void ConnectWMI()

// Creates command prompt process and runs the command provided as param
void LaunchRemoteProcess()
```


####Public Events
```c#
event EventHandler OutputUpdated	// Fires each time the output is updated; including connection logging
event EventHandler CmdFinished		// Fires after the command has completed and output has been collected
```

========================
NetConnect Utility Class
------------------------
Uses Windows API to make connection to network resources
- WNetUseConnection
- WNetCancelConnection2

```c#
// Makes connection to UNC path specified and returns Error string
string ConnectNetResource(string remoteUNC, string username, string password)

// Disconnects mapped network resource at specified UNC path and returns Error string
string DisconnectNetResource(string remoteUNC)
```
