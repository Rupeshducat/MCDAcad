using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

using MCD;
namespace MCD
{
    static class Program
    {
        

        [DllImport("user32", EntryPoint = "SendMessageA")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg,int wParam, COPYDATASTRUCT lParam);

        
        public static string ProcessWM_COPYDATA(System.Windows.Forms.Message m)
        {
            if (m.WParam.ToInt32() == FunctionsNvar._messageID )
            {
                COPYDATASTRUCT st =(COPYDATASTRUCT)Marshal.PtrToStructure(m.LParam,typeof(COPYDATASTRUCT));
                return st.lpData;
            }
            return null;
        }

        /// <summary>
        /// Structure required to be sent with the WM_COPYDATA message
        /// This structure is used to contain the CommandLine
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public class COPYDATASTRUCT
        {
            public int dwData = 0;//32 bit int to passed. Not used.
            public int cbData = 0;//length of string. Will be one greater because of null termination. 
            public string lpData;//string to be passed.

            public COPYDATASTRUCT()
            {
            }

            public COPYDATASTRUCT(string Data)
            {
                lpData = Data + "\0"; //add null termination
                cbData = lpData.Length; //length includes null chr sowill be one greater
            }
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            bool createdNew;
            string process_name =
            Process.GetCurrentProcess().ProcessName.ToString();
            Mutex m = new Mutex(true, process_name, out createdNew);
            
            if (!createdNew)
            {
                //IntPtr hWnd = GetHWndOfPrevInstance(Process.GetCurrentProcess().ProcessName);
                //IntPtr hWnd = GetHWndOfPrevInstance(Process.GetCurrentProcess().ProcessName);
                //SendMessage(hWnd, FunctionsNvar.WM_COPYDATA, FunctionsNvar._messageID, new COPYDATASTRUCT(Environment.CommandLine));
                
                return;
            }

            //if (args.Length == 3)
            //{
            //    FunctionsNvar.FilePath = args[0];
            //    FunctionsNvar.AppId = args[1];
            //    FunctionsNvar.DbStatus = args[2];
            //    if (string.IsNullOrEmpty(args[0]) == false && System.IO.File.Exists(FunctionsNvar.FilePath) == true 
            //        && string.IsNullOrEmpty(args[1]) == false && string.IsNullOrEmpty(args[2]) == false )
            //    {
            //        MessageBox.Show("HI");
            //    }
                
            //}
            //else if (args.Length == 0)
            //{
             
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainFrm());
            }
            catch (Exception ex)
            {
                //string msg = "<p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:20.0pt;color:#C00000'>" +
                //               "Error Occured " +
                //               "</span><span style='color:#C00000'><o:p></o:p></span></b></p>" +
                //               "<p class=MsoNormal><b><span style='color:#1F497D'>" + ex.Message + "</span></b></p>" +
                //               "<p class=MsoNormal><b><span style='color:#1F497D'>" + ex.StackTrace + "</span></b></p>";
                //MailToMCD.sendMail(msg);
                MessageBox.Show(ex.Message + ex.StackTrace );
            }
                
                

            //}
            GC.KeepAlive(m);
            
        }

        /*
        /// <summary>
        /// Searches for a previous instance of this app.
        /// </summary>
        /// <returns>
        /// hWnd of the main window of the previous instance
        /// or IntPtr.Zero if not found.
        /// </returns>
        private static IntPtr GetHWndOfPrevInstance(string ProcessName)
        {
            //get the current process
            Process CurrentProcess = Process.GetCurrentProcess();
            //get a collection of the currently active processes withthe same name
            Process[] Ps = Process.GetProcessesByName(ProcessName);
            //if only one exists then there is no previous instance
            if (Ps.Length >= 1)
            {
                foreach (Process P in Ps)
                {
                    if (P.Id != CurrentProcess.Id)//ignore thisprocess
                    {
                        //weed out apps that have the same exe namebut are started from a different filename.
                        if (P.ProcessName == ProcessName)
                        {
                            IntPtr hWnd = IntPtr.Zero;
                            try
                            {
                                //if process does not have a MainWindowHandle then an exception will be thrown
                                //so catch and ignore the error.
                                hWnd = P.MainWindowHandle;
                            }
                            catch { }
                            //return if hWnd found.
                            if (hWnd.ToInt32() != 0) return hWnd;
                        }
                    }
                }
            }
            return IntPtr.Zero;
        }
        */

    }
}
