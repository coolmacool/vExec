vExec
=====

proof of concept

Initial work on a C# WMI port for interactive remote command execution

Allows running of net use command with network resources available to logged on user 


Example
--------------------


```c#
//Instantiation
vExec ve = new vExec("MACHINEADDRESS", "username", "password");

//Delegate for updating UI from output received from OutputUpdated event
delegate void OutputDelegate(string output);

//Subscribe to OutputUpdated Event
ve.OutputUpdated += new EventHandler(ve_OutputUpdated);

private void ve_OutputUpdated(object sender, EventArgs e)
{
	SetOutput(ve.CmdOutput);
}

//Method to update UI
private void SetOutput(string output)
{
	if (InvokeRequired)
	{
		Invoke(new OutputDelegate(SetOutput), new object[] { output });
	}
	else
	{
		txtOutput.Text = output;
	}
}

//Method to begin remote command execution
private void button1_Click(object sender, EventArgs e)
{      
	ve.Run("ipconfig");
}
```
For WPF Applications
---------------------
Replace Invoke & Invoke with Dispatcher.CheckAccess() & Dispatcher.Invoke

```c#
if (!Dispatcher.CheckAccess())
{
	Dispatcher.Invoke(new outputDelegate(SetOutput), new object[] { output });
}
else
{
	txtOutput.Text = output;
}

```
