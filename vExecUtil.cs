using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq;

namespace vExecUtil
{
    public class vNetConnect
    {
        // Add Win32 API Platform Invoke functionality
        #region Consts
        const int RESOURCE_CONNECTED = 0x00000001;
        const int RESOURCE_GLOBALNET = 0x00000002;
        const int RESOURCE_REMEMBERED = 0x00000003;

        const int RESOURCETYPE_ANY = 0x00000000;
        const int RESOURCETYPE_DISK = 0x00000001;
        const int RESOURCETYPE_PRINT = 0x00000002;

        const int RESOURCEDISPLAYTYPE_GENERIC = 0x00000000;
        const int RESOURCEDISPLAYTYPE_DOMAIN = 0x00000001;
        const int RESOURCEDISPLAYTYPE_SERVER = 0x00000002;
        const int RESOURCEDISPLAYTYPE_SHARE = 0x00000003;
        const int RESOURCEDISPLAYTYPE_FILE = 0x00000004;
        const int RESOURCEDISPLAYTYPE_GROUP = 0x00000005;

        const int RESOURCEUSAGE_CONNECTABLE = 0x00000001;
        const int RESOURCEUSAGE_CONTAINER = 0x00000002;


        const int CONNECT_INTERACTIVE = 0x00000008;
        const int CONNECT_PROMPT = 0x00000010;
        const int CONNECT_REDIRECT = 0x00000080;
        const int CONNECT_UPDATE_PROFILE = 0x00000001;
        const int CONNECT_COMMANDLINE = 0x00000800;
        const int CONNECT_CMD_SAVECRED = 0x00001000;

        const int CONNECT_LOCALDRIVE = 0x00000100;
        #endregion

        #region Errors
        const int NO_ERROR = 0;

        const int ERROR_ACCESS_DENIED = 5;
        const int ERROR_ALREADY_ASSIGNED = 85;
        const int ERROR_BAD_DEVICE = 1200;
        const int ERROR_BAD_NET_NAME = 67;
        const int ERROR_BAD_PROVIDER = 1204;
        const int ERROR_CANCELLED = 1223;
        const int ERROR_EXTENDED_ERROR = 1208;
        const int ERROR_INVALID_ADDRESS = 487;
        const int ERROR_INVALID_PARAMETER = 87;
        const int ERROR_INVALID_PASSWORD = 1216;
        const int ERROR_MORE_DATA = 234;
        const int ERROR_NO_MORE_ITEMS = 259;
        const int ERROR_NO_NET_OR_BAD_PATH = 1203;
        const int ERROR_NO_NETWORK = 1222;

        const int ERROR_BAD_PROFILE = 1206;
        const int ERROR_CANNOT_OPEN_PROFILE = 1205;
        const int ERROR_DEVICE_IN_USE = 2404;
        const int ERROR_NOT_CONNECTED = 2250;
        const int ERROR_OPEN_FILES = 2401;

        private struct ErrorClass
        {
            public int num;
            public string message;
            public ErrorClass(int num, string message)
            {
                this.num = num;
                this.message = message;
            }
        }


        // Created with excel formula:
        // ="new ErrorClass("&A1&", """&PROPER(SUBSTITUTE(MID(A1,7,LEN(A1)-6), "_", " "))&"""), "
        private static ErrorClass[] ERROR_LIST = new ErrorClass[] {
			new ErrorClass(ERROR_ACCESS_DENIED, "Error: Access Denied"), 
			new ErrorClass(ERROR_ALREADY_ASSIGNED, "Error: Already Assigned"), 
			new ErrorClass(ERROR_BAD_DEVICE, "Error: Bad Device"), 
			new ErrorClass(ERROR_BAD_NET_NAME, "Error: Bad Net Name"), 
			new ErrorClass(ERROR_BAD_PROVIDER, "Error: Bad Provider"), 
			new ErrorClass(ERROR_CANCELLED, "Error: Cancelled"), 
			new ErrorClass(ERROR_EXTENDED_ERROR, "Error: Extended Error"), 
			new ErrorClass(ERROR_INVALID_ADDRESS, "Error: Invalid Address"), 
			new ErrorClass(ERROR_INVALID_PARAMETER, "Error: Invalid Parameter"), 
			new ErrorClass(ERROR_INVALID_PASSWORD, "Error: Invalid Password"), 
			new ErrorClass(ERROR_MORE_DATA, "Error: More Data"), 
			new ErrorClass(ERROR_NO_MORE_ITEMS, "Error: No More Items"), 
			new ErrorClass(ERROR_NO_NET_OR_BAD_PATH, "Error: No Net Or Bad Path"), 
			new ErrorClass(ERROR_NO_NETWORK, "Error: No Network"), 
			new ErrorClass(ERROR_BAD_PROFILE, "Error: Bad Profile"), 
			new ErrorClass(ERROR_CANNOT_OPEN_PROFILE, "Error: Cannot Open Profile"), 
			new ErrorClass(ERROR_DEVICE_IN_USE, "Error: Device In Use"), 
			new ErrorClass(ERROR_EXTENDED_ERROR, "Error: Extended Error"), 
			new ErrorClass(ERROR_NOT_CONNECTED, "Error: Not Connected"), 
			new ErrorClass(ERROR_OPEN_FILES, "Error: Open Files"), 
		};

        private static string getErrorForNumber(int errNum)
        {
            foreach (ErrorClass er in ERROR_LIST)
            {
                if (er.num == errNum) return er.message;
            }
            return "Error: Unknown, " + errNum;
        }
        #endregion

        #region Win32 API
        [DllImport("Mpr.dll")]
        private static extern int WNetUseConnection(
            IntPtr hwndOwner,
            NETRESOURCE lpNetResource,
            string lpPassword,
            string lpUserID,
            int dwFlags,
            string lpAccessName,
            string lpBufferSize,
            string lpResult
            );

        [DllImport("Mpr.dll")]
        private static extern int WNetCancelConnection2(
            string lpName,
            int dwFlags,
            bool fForce
            );

        [StructLayout(LayoutKind.Sequential)]
        private class NETRESOURCE
        {
            public int dwScope = 0;
            public int dwType = 0;
            public int dwDisplayType = 0;
            public int dwUsage = 0;
            public string lpLocalName = "";
            public string lpRemoteName = "";
            public string lpComment = "";
            public string lpProvider = "";
        }
        #endregion

        #region Methods

        public static string ConnectNetResource(string remoteUNC, string username, string password)
        {
            NETRESOURCE nr = new NETRESOURCE();
            nr.dwType = RESOURCETYPE_DISK;
            nr.lpRemoteName = remoteUNC;

            int ret;

            ret = WNetUseConnection(IntPtr.Zero, nr, password, username, 0, null, null, null);

            if (ret == NO_ERROR)
            {
                isConnected = true;
                return null;
            }
            return getErrorForNumber(ret);
        }

        public static string DisconnectNetResource(string remoteUNC)
        {
            int ret = WNetCancelConnection2(remoteUNC, CONNECT_UPDATE_PROFILE, false);
            if (ret == NO_ERROR)
            {
                isConnected = false;
                return null;
            }
            return getErrorForNumber(ret);
        }

        #endregion

        public static bool isConnected { get; set; }
    }
    /*
    public class OutputChangedEventArgs : EventArgs
    {
        public string CmdOutput { get; set; }
    }
    */
    public class vExec
    {
        /**************
         * CONSTANTS
         *************/
        private const string TempFileLocation = @"C:\temp";
        //private const string JobTempFileName = "vexec_job.tmp";
        private const string CmdTempFileName = "vexec_cmd.tmp";
        //private const string JobTempFile = TempFileLocation + @"\" + JobTempFileName;
        private const string CmdTempFile = TempFileLocation + @"\" + CmdTempFileName;
        

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


        /**************
        * CONSTRUCTORS
        **************/
        public vExec() { }

        public vExec(string remoteComputer, string cmd)
        {
            RemoteComputer = remoteComputer;
            RemoteCommand = cmd;
            UserName = UserPass = String.Empty;
        }

        public vExec(string remoteComputer, string cmd, string userName)
        {
            RemoteComputer = remoteComputer;
            RemoteCommand = cmd;
            UserName = userName;
            UserPass = String.Empty;
        }

        public vExec(string remoteComputer, string cmd, string userName, string userPass)
        {
            RemoteComputer = remoteComputer;
            RemoteCommand = cmd;
            UserName = userName;
            UserPass = userPass;
        }

        
        
        /**************
        * DELEGATES
        **************/
        public delegate void OutputCallback(string output);

        protected virtual void OnOutputCreated(object sender, FileSystemEventArgs e)
        {
            if (e.Name == CmdTempFileName)
            {
                using (StreamReader sr = new StreamReader(_remotePath + @"\" + CmdTempFileName))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        _cmdOutput += line + "\r\n";
                    }

                }
                if (OutputChanged != null)
                {
                    OutputChanged(this, e);
                }
                CleanUp();
                
                //PrepareOutput();
            }
        }
        /*
        private void OutputPreparedEvent(OutputChangedEventArgs e)
        {
            EventHandler<OutputChangedEventArgs> handler = OutputChanged;
            if (handler != null)
                handler(this, e);
        }

        private void PrepareOutput()
        {
            OutputChangedEventArgs args = new OutputChangedEventArgs();
            args.CmdOutput = _cmdOutput;
            OutputPreparedEvent(args);
        }
        */
        /**************
        * EVENTS
        **************/
        public event EventHandler OutputChanged;
        //public event EventHandler<OutputChangedEventArgs> OutputChanged;

        /**************
        * PROPERTIES
        **************/
        public string RemoteComputer { get; set; }
        public string RemoteCommand { get; set; }
        public string UserName { get; set; }
        public string UserPass { get; set; }
        public string CmdOutput 
        { 
            get { return _cmdOutput; } 
        }
        

        /**************
        * METHODS
        **************/
        public void ConnectWMI()
        {
            try
            {
                // Initialize WMI Connection Options
                _conOptions = new ConnectionOptions()
                {
                    Username = UserName,
                    Password = UserPass,
                    Impersonation = ImpersonationLevel.Impersonate, // Allow WMI to run under specified credentials
                    Authentication = AuthenticationLevel.PacketPrivacy  // Encrypt entire communication between the client and server
                };

                // Connect to the WMI Management script namespace
                _scope = new ManagementScope(@"\\" + RemoteComputer + @"\root\cimv2", _conOptions);
                _scope.Connect();

                // Make connection to remote IPC$ share to allow FileSystemWatcher to monitor for temporary file creation/changes
                vNetConnect.ConnectNetResource(@"\\" + RemoteComputer + @"\IPC$", UserName, UserPass);
                _remotePath = @"\\" + RemoteComputer + @"\C$\temp";
                _watch = new FileSystemWatcher()
                {
                    Path = _remotePath,
                    Filter = "*.tmp",
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.CreationTime | NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };
                _watch.Changed += new FileSystemEventHandler(OnOutputCreated);
            }
            catch(Exception e)
            {
                
            }
        }

        public void LaunchRemoteProcess()
        {
            try
            {
                // Bind scope to ManagementClass using Win32_Process WMI Class
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

                _inParams["CommandLine"] = "CMD /C " +
                                                "SCHTASKS /Create /TN \"vexectemp\" /TR \"CMD /C " + RemoteCommand + " > " + CmdTempFile + "\" /SC MONTHLY " + "/RU " + UserName + " /RP " + UserPass +
                                                " & PING 192.0.2.2 -n 1 -w 500 >NUL & " +
                                                "SCHTASKS /Run /TN \"vexectemp\"" + " & PING 192.0.2.2 -n 1 -w 500 >NUL & " +
                                                "SCHTASKS /Delete /TN \"vexectemp\" /F";
                _cmdOutput = "Sending remote cmd...\r\n";  
                _outParams = _mgmt.InvokeMethod("Create", _inParams, null);

                if (_outParams["returnValue"].ToString() != "0")
                {
                    _cmdOutput = "Failed to send remote command\r\nError Level: " + _outParams["returnValue"] + "\r\n\r\n";
                }
                else
                {
                    _cmdOutput += "Command sent successfully.\r\n\r\n";
                }

            }
            catch (Exception e)
            {

            }
        }

        private void CleanUp()
        {
            //Thread.Sleep(1000);
            _mgmt.Dispose();
            if (File.Exists(_remotePath + @"\" + CmdTempFileName))
                File.Delete(_remotePath + @"\" + CmdTempFileName);
            //if (File.Exists(_remotePath + @"\" + JobTempFileName))
            //    File.Delete(_remotePath + @"\" + JobTempFileName);
            //vNetConnect.DisconnectNetResource(@"\\" + RemoteComputer + @"\IPC$");

        }
    }
}
