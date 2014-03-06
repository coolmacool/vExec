namespace vExecUtil
{
    using System;
    using System.ComponentModel;
    using System.IO;
    using System.Management;
    using System.Runtime.InteropServices;
    using System.Threading;

    public class vExec
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int MessageBox(IntPtr hWnd, String text, String caption, uint type);

        /**************
         * CONSTANTS
         *************/
        private const string TempFileLocation = @"C:\temp";
        private const string CmdTempFileName = "vexec_cmd.tmp";
        private const string CmdTempFile = TempFileLocation + @"\" + CmdTempFileName;
        private const string CmdPause = " & PING 192.0.2.2 -n 1 -w 500 >NUL & ";
        

        /**************
        * FIELDS
        **************/
        private string _remotePath;
        private string _cmdOutput;
        private ConnectionOptions _conOptions;
        private ManagementScope _scope;
        private ManagementClass _mgmt;
        private ManagementBaseObject _inParams;
        private ManagementBaseObject _outParams;
        private FileSystemWatcher _watch;
        private BackgroundWorker _worker;

        /**************
        * CONSTRUCTORS
        **************/
        public vExec() { }

        public vExec(string remoteComputer)
        {
            RemoteComputer = remoteComputer;
            UserName = UserPass = String.Empty;
        }

        public vExec(string remoteComputer, string userName)
        {
            RemoteComputer = remoteComputer;
            UserName = userName;
            UserPass = String.Empty;
        }

        public vExec(string remoteComputer, string userName, string userPass)
        {
            RemoteComputer = remoteComputer;
            UserName = userName;
            UserPass = userPass;
        }


        /**************
        * DELEGATES
        **************/
        // Delegate for command output file creation event to read file
        protected virtual void OnOutputCreated(object sender, FileSystemEventArgs e)
        {
            try
            {
                if (e.Name == CmdTempFileName)
                {
                    string buffer = "";
                    using (StreamReader sr = new StreamReader(_remotePath + @"\" + CmdTempFileName))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            buffer += line + "\r\n";
                        }
                    }

                    UpdateOutput(buffer);
                    OnCmdFinished(EventArgs.Empty);
                    CleanUp();
                }
            }
            catch(Exception ex)
            {
                MessageBox(IntPtr.Zero, "An Error has occurred.\r\n" + ex.Message, "Error", 0);
                return;
            }
        }

        // Notify output has been updated
        protected void OnOutputUpdated(EventArgs e)
        {
            if (OutputUpdated != null)
                OutputUpdated(this, e);
        }

        // Notify command has finished
        protected void OnCmdFinished(EventArgs e)
        {
            if (CmdFinished != null)
                CmdFinished(this, e);
        }

        /**************
        * EVENTS
        **************/
        public event EventHandler OutputUpdated;
        public event EventHandler CmdFinished;


        /**************
        * PROPERTIES
        **************/
        public string RemoteComputer { get; set; }
        public string RemoteCommand { get; set; }
        public string UserName { get; set; }
        public string UserPass { get; set; }
        public string CmdOutput { get { return _cmdOutput; } }
        

        /**************
        * METHODS
        **************/
        public void ConnectWMI()
        {
            try
            {
                UpdateOutput("Connecting to remote machine...", true);

                // Initialize WMI Connection Options
                _conOptions = new ConnectionOptions()
                {
                    Username = UserName,
                    Password = UserPass,
                    Impersonation = ImpersonationLevel.Impersonate, // Allow WMI to run under specified credentials
                    Authentication = AuthenticationLevel.PacketPrivacy  // Encrypt communication between machines
                };

                // Connect to the WMI Management script namespace
                _scope = new ManagementScope(@"\\" + RemoteComputer + @"\root\cimv2", _conOptions);
                _scope.Connect();

                // Make connection to remote IPC$ share to allow FileSystemWatcher to monitor for temporary file creation/changes
                NetConnect.ConnectNetResource(@"\\" + RemoteComputer + @"\IPC$", UserName, UserPass);
                _remotePath = @"\\" + RemoteComputer + @"\C$\temp";

                // Create a FileSystemWatcher to monitor for command output file creation
                _watch = new FileSystemWatcher()
                {
                    Path = _remotePath,
                    Filter = "*.tmp",
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.CreationTime | NotifyFilters.LastWrite | NotifyFilters.LastAccess,
                    EnableRaisingEvents = true
                };
                _watch.Changed += new FileSystemEventHandler(OnOutputCreated);

                UpdateOutput("Connected successfully.");
            }
            catch(Exception e)
            {
                throw new ApplicationException(e.Message);
            }
        }

        public void LaunchRemoteProcess()
        {
            try
            {
                UpdateOutput("Sending remote command...");

                // Create a new Win32_Process ManagementClass
                _mgmt = new ManagementClass(_scope, new ManagementPath("Win32_Process"), new ObjectGetOptions());

                _inParams = _mgmt.GetMethodParameters("Create");

                /* ******************************************************************************
                 * 
                 * Create process for remote command
                 *  
                 *  1) Command is created as a Monthly Scheduled Task with specified credentials
                 *  2) Scheduled Task is submitted to run immediately
                 *  3) Scheduled Task is deleted after a 500 millisecond pause
                 *  
                 *******************************************************************************/

                /* Can be run using a hidden window using the SYSTEM account, however any mapped network resources would be unavailable to logged on user
                 * 
                 * _inParams["CommandLine"] = "CMD /C " +
                                                "SCHTASKS /Create /TN \"vexectemp\" /TR \"CMD /C " + RemoteCommand + " > " + CmdTempFile + " 2>&1\" /SC MONTHLY /RU \"SYSTEM\"" +
                                                CmdPause +
                                                "SCHTASKS /Run /TN \"vexectemp\"" +
                                                CmdPause +
                                                "SCHTASKS /Delete /TN \"vexectemp\" /F";
                 */

                _inParams["CommandLine"] = "CMD /C " +
                                                "SCHTASKS /Create /TN \"vexectemp\" /TR \"CMD /C " + RemoteCommand + " > " + CmdTempFile + " 2>&1\" /SC MONTHLY /RU " + UserName + " /RP " + UserPass +
                                                CmdPause +
                                                "SCHTASKS /Run /TN \"vexectemp\"" + 
                                                CmdPause +
                                                "SCHTASKS /Delete /TN \"vexectemp\" /F";

                // Create process and collect result
                _outParams = _mgmt.InvokeMethod("Create", _inParams, null);

                if (_outParams["returnValue"].ToString() != "0")
                {
                    UpdateOutput("Command failed!\r\nError Level: " + _outParams["returnValue"] + "\r\n");
                }
                else
                {
                    UpdateOutput("Command sent successfully.");
                }

            }
            catch (Exception e)
            {
                throw new ApplicationException(e.Message);
            }
        }

        public void Run(string command)
        {
            try
            {
                RemoteCommand = command;
                _worker = new BackgroundWorker();
                _worker.DoWork += new DoWorkEventHandler(worker_DoWork);
                _worker.RunWorkerAsync();
            }
            catch (Exception e)
            {
                MessageBox(IntPtr.Zero, "An Error has occurred.\r\n" + e.Message, "Error", 0);
                return;
            }
        }

        private void CleanUp()
        {
            try
            {
                // Cleanup objects and temporary files
                _mgmt.Dispose();
                _watch.Dispose();
                _worker.Dispose();

                if (File.Exists(_remotePath + @"\" + CmdTempFileName))
                    File.Delete(_remotePath + @"\" + CmdTempFileName);

                Thread.Sleep(2000);

                NetConnect.DisconnectNetResource(@"\\" + RemoteComputer + @"\IPC$");
            }
            catch (Exception e)
            {
                MessageBox(IntPtr.Zero, "An Error has occurred during cleanup.\r\n" + e.Message, "Error", 0);
                return;
            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                ConnectWMI();
                LaunchRemoteProcess();
            }
            catch (Exception ex)
            {
                MessageBox(IntPtr.Zero, "An Error has occurred.\r\n" + ex.Message, "Error", 0);
                return;
            }
        }

        private void UpdateOutput(string updatedOutput, bool resetOutput = false)
        {
            if (resetOutput)
            {
                _cmdOutput = "";
            }

            // Update the output and raise notify event
            _cmdOutput += updatedOutput + "\r\n";
            OnOutputUpdated(EventArgs.Empty);
        }
    }
}
