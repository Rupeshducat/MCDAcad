using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Collections;
using System.Reflection;
using System.IO;
using Microsoft.Win32;
using IBM.Data.DB2;
using log4net;

using Autodesk.AutoCAD.Interop.Common;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using MCD;
using Autodesk.AutoCAD.Internal;

using System.Configuration;

[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]
namespace MCD
{
    public partial class MainFrm : Form
    {
        //private static readonly ILog log = LogManager.GetLogger(typeof(TestPage1).Name);  
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        [System.Runtime.InteropServices.DllImport("user32")]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
        private const int WM_ACTIVATEAPP = 0x001C;
        //private bool appActive = true;
        private List<Thread> Threads = new List<Thread>();
        private DateTime now = DateTime.Now;
        private byte count = 0;
        private DB2DataReader MainReader;
        public static int appstage;
        public MainFrm()
        {

            InitializeComponent();

        }

        //[System.Security.Permissions.PermissionSet(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
        //protected override void WndProc(ref Message m)
        //{
        //    if (m.Msg == FunctionsNvar.WM_COPYDATA)
        //    {
        //        string command = Program.ProcessWM_COPYDATA(m);
        //        if (command != null)
        //        {
        //            string[] filepath = command.Split(' ');
        //            FunctionsNvar.FilePath = filepath[1];
        //            Thread th = new Thread(new ThreadStart(Openacad));
        //            th.Start();
        //            //th.Join();

        //            //processCommandLine(command);
        //            //return;
        //        }
        //    }
        //    base.WndProc(ref m);
        //}

        public static AcadApplication StartAutoCADSession()
        {
            // Each time create a new instance of AutoCAD


            const string progID = "AutoCAD.Application.24.2";

            AcadApplication acApp = null;
            try
            {
                log.Debug("StartAutoCADSession() - Started");
                Type acType = Type.GetTypeFromProgID(progID);
                acApp = (AcadApplication)Activator.CreateInstance(acType, true);
                log.Debug("StartAutoCADSession() - Ended");
            }
            catch (System.Exception ex)
            {
                //Environment.Exit(0); 
                log.Error("StartAutoCADSession()-Unable to start Autocad session-Error(" + ex.Message + ")");
                //MessageBox.Show("Error Occured : " + ex.Message + "\n" + ex.StackTrace);
            }

            return acApp;
        }

        protected class opencadclass
        {
            string Path;
            public string ID_ver;
            public opencadclass(string pathname, string id_Ver)
            {
                log.Debug("opencadclass():Starts with (" + id_Ver + ") and (" + pathname + ")");
                Path = pathname;
                ID_ver = id_Ver;
                log.Debug("opencadclass():Ends(" + id_Ver + ") and (" + pathname + ")");
            }


            public void Openacad()
            {
                try
                {
                    log.Debug("Openacad() - Started");
                    AcadApplication acapp = StartAutoCADSession();
                    System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select APP_STAGE from application where ID_VER = " + ID_ver + ";commit;");
                    appstage = int.Parse(MainDtApp.Rows[0][0].ToString());
                    //-->>> Included getting Appstage of Application ID into log while selecting it on 10th July 2013 By Kiran Bishaj.
                    // AppStageslog.DebugLog("Openacad()- Selected Application ID " + ID_ver + " with APP_STAGE (" + appstage + ")");
                    //<<<-- Included getting Appstage of Application ID into log while selecting it on 10th July 2013 By Kiran Bishaj.
                    // IntPtr hnwdintptr = (IntPtr)acapp.HWND;
                    if (acapp == null)
                    {
                        if (appstage == 23 || appstage == 24)
                        {

                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = " + appstage + " where ID_VER = " + ID_ver + ";commit;");
                            //-->>> Included Appstage log on 29-05-2013 By Kiran  
                            AppStageslog.DebugLog("Openacad()- Updated APP_STAGE  is (" + appstage + ") for Application ID " + ID_ver);
                            //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                        }
                        else
                        {
                            //Updating APP_STAGE_TEMP to null to avoid strucking of request id's at APP_STAGE_TEMP 4 on 23rd Sept 2013.
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 1,APP_STAGE_TEMP = NULL where ID_VER = " + ID_ver + ";commit;");
                            //-->>> Included Appstage log on 29-05-2013 By Kiran  
                            AppStageslog.DebugLog("Openacad()- Updated APP_STAGE is 1 its APP_STAGE_TEMP is NULL for Application ID " + ID_ver);
                            //<<<-- Included Appstage log on 29-05-2013 By Kiran
                        }



                        return;

                    }
                    else
                    {
                        IntPtr hnwdintptr = (IntPtr)acapp.HWND;
                        try
                        {
                            log.Debug("Openacad()-Inserting into EXCEPTIONREMARKS table");
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";INSERT INTO EXCEPTIONREMARKS(EXCPTREMRKS_ID_VER, EXCPTREMRKS_REMARKS)" +
                                "VALUES (" + ID_ver + "," + acapp.HWND.ToString() + ");");
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set DRAWING_TIME = '" + DateTime.Now.ToShortTimeString() + "' where ID_VER = " + ID_ver + ";commit;");
                            log.Debug("Openacad()-Updated DRAWING_TIME in DRAWING table");

                            AcadDocument acd;

                            try
                            {
                                acd = acapp.Documents.Open(Path, false, "");

                            }
                            catch (System.Exception ex)
                            {
                                log.Error("Openacad()-Unable to Open Autocad -Error(" + ex.Message + ")");
                                string id = ID_ver.Substring(0, ID_ver.Length - 2);
                                string ver = ID_ver.Substring(ID_ver.Length - 2);
                                string TxtFilePath = @"d:\mcd\Report\" + id + "_" + ver + "_ValidationReport.Txt";
                                using (StreamWriter TxtFile = new StreamWriter(TxtFilePath, true))
                                {
                                    TxtFile.WriteLine("Drawing can not open, Please upload valid dwg");
                                }
                                ValidateReport vr = new ValidateReport();
                                vr.validReport(ID_ver, TxtFilePath);
                                System.IO.File.Delete(TxtFilePath);
                                if ((appstage == 23) || (appstage == 24))
                                {
                                }
                                else
                                {
                                    FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 51,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_ver + ";commit;");
                                    //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                    AppStageslog.DebugLog("Openacad()-Updated APP_STAGE is 51 and its APP_STAGE_TEMP is NULL for Application ID " + ID_ver);
                                    //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                }

                                DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
                                con.Open();
                                DB2Command Cmd1 = new DB2Command("set schema " + FunctionsNvar.schema + ";select F_ID from application where id_ver =" + ID_ver + ";commit;", con);
                                int fid = Convert.ToInt16(Cmd1.ExecuteScalar());
                                StringBuilder Validation_FileName = new StringBuilder();
                                switch (fid)
                                {
                                    case 1:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport.PDF");
                                        break;
                                    case 2:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_CC.PDF");
                                        break;
                                    case 3:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised.PDF");
                                        break;
                                    case 4:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Regularized.PDF");
                                        break;
                                    case 5:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_AA.PDF");
                                        break;
                                    case 6:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_REVDN.PDF");
                                        break;
                                    case 7:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SARAL_Revise.PDF");
                                        break;
                                    case 8:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SANCTION_Up_To_500_Sqmt.PDF");
                                        break;
                                    case 9:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised_SANCTION_Up_To_500_Sqmt.PDF");
                                        break;
                                }
                                con.Close();

                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set REPORT_FILE_NAME = '" + Validation_FileName + "' where ID_VER = " + id + ver + ";commit;");
                                return;

                            }

                            if (appstage == 24 || appstage == 23)
                            {

                            }
                            else
                            {

                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 45 where ID_VER = " + ID_ver + ";commit;");
                                //-->>> Included Appstage log on 29-05-2013 By Kiran   
                                AppStageslog.DebugLog("Openacad()-Updated APP_STAGE is 45  for Application ID " + ID_ver);
                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                            }
                            try
                            {
                                acapp.Visible = true;
                            }
                            catch (System.Exception)
                            {

                            }
                            try
                            {
                                acd.SetVariable("autosnap", 63);
                            }
                            catch (System.Exception)
                            {

                            }



                            Thread.Sleep(10000);
                            try
                            {
                                acd.Close(false, Path);
                            }
                            catch (System.Exception)
                            {

                            }

                            finally
                            {

                            }


                        }
                        catch (System.Exception ex)
                        {
                            log.Error("Openacad()-Unable to Open Autocad -Error(" + ex.Message + ")");
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 44,APP_STAGE_TEMP = NULL where ID_VER = " + ID_ver + ";commit;");
                            //MessageBox.Show("Error Occured : " + ex.Message + "\n" + ex.StackTrace);
                            //-->>> Included Appstage log on 29-05-2013 By Kiran  
                            AppStageslog.DebugLog("Openacad()-Updated APP_STAGE is 44 and its APP_STAGE_TEMP is NULL for Application ID " + ID_ver);
                            //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                        }
                        finally
                        {
                            //PlotDwg.PlotCurrentLayout(
                            try
                            {
                                int processid;
                                //IntPtr hnwdintptr = (IntPtr)acapp.HWND;
                                int threadid = GetWindowThreadProcessId(hnwdintptr, out processid);
                                System.Diagnostics.Process Pracad = System.Diagnostics.Process.GetProcessById(processid);

                                Pracad.Kill();
                                //acapp.Quit();
                            }
                            catch (System.Exception)
                            {

                            }

                            if (System.IO.File.Exists(Path) == true)
                            {
                                try
                                {
                                    System.IO.File.Delete(Path);
                                }
                                catch (System.Exception)
                                {

                                }
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {

                    MessageBox.Show("Error Occured : " + ex.Message + "\n" + ex.StackTrace);
                }
            }

        }

        //private void ChkDbTimer_Tick(object sender, EventArgs e)
        //{ string dwgName = String.Empty;
        //}
        private void ChkDbTimer_Tick(object sender, EventArgs e)
        {
            string dwgName = String.Empty;
            string id_no = String.Empty;
            string id = String.Empty;
            string ver = String.Empty;
            string TxtFilePath = String.Empty;
            int Appstage, fid;
            TimeSpan ts = DateTime.Now.Subtract(now);
            StringBuilder Validation_FileName = new StringBuilder();
            System.Data.DataTable FileDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application where (APP_STAGE = 2 or APP_STAGE = 3 or APP_STAGE = 1) and APP_STAGE_TEMP IS NULL;commit;");
            //System.Data.DataTable FileDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application a,SYSTEM_NUMBER s where a.ID_VER=s.ID_VER and (a.APP_STAGE = 2 or a.APP_STAGE = 3 or a.APP_STAGE = 1) and APP_STAGE_TEMP IS NULL and s.SYS_N0=7;commit;");   
            for (int i = 0; i < FileDtApp.Rows.Count; i++)
            {
                string ID_VER = FileDtApp.Rows[i]["id_ver"].ToString();
                Int16 appstage = (Int16)FileDtApp.Rows[i]["APP_STAGE"];
                //-->>> Included getting Appstage of Application ID into log while selecting it on 10th July 2013 By Kiran Bishaj.
                // AppStageslog.DebugLog("ChkDbTimer_Tick()- Selected Application ID " + ID_VER + " with APP_STAGE (" + appstage + ")");
                //<<<-- Included getting Appstage of Application ID into log while selecting it on 10th July 2013 By Kiran Bishaj.

                System.Data.DataTable DtDwg = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from drawing where ID_VER = " + ID_VER + " order by  DWG_VER DESC;");
                if (DtDwg.Rows.Count == 0)
                {
                    //FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE_TEMP = 7 where ID_VER = " + ID + ";commit;");
                    continue;
                }
                dwgName = DtDwg.Rows[0]["DWGNAME"].ToString();
                id_no = DtDwg.Rows[0]["ID"].ToString();
                string[] dwgId = dwgName.Split('_');
                if ((String.IsNullOrEmpty(dwgName)) || (dwgId[0].ToString() != id_no.ToString()))
                {
                    id = ID_VER.Substring(0, ID_VER.Length - 2);
                    ver = ID_VER.Substring(ID_VER.Length - 2);
                    TxtFilePath = @"d:\mcd\Report\" + id + "_" + ver + "_ValidationReport.Txt";
                    using (StreamWriter TxtFile = new StreamWriter(TxtFilePath, true))
                    {
                        TxtFile.WriteLine(MCD.ConstantStrings.STR_DWGNAME_ISSUE_TXT);
                    }
                    ValidateReport vr = new ValidateReport();
                    vr.validReport(ID_VER, TxtFilePath);
                    System.IO.File.Delete(TxtFilePath);
                    Appstage = AppnDbquery(ID_VER);
                    if ((MCD.ConstantStrings.INT_FILE_PROCESS_START1 != Appstage) || (MCD.ConstantStrings.INT_FILE_PROCESS_START2 != Appstage))
                    {
                        FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 51,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_VER + ";commit;");
                        AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 51 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                    }
                    fid = AppnDbquery(ID_VER);
                    switch (fid)
                    {
                        case 1:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport.PDF");
                            break;
                        case 2:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_CC.PDF");
                            break;
                        case 3:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised.PDF");
                            break;
                        case 4:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Regularized.PDF");
                            break;
                        case 5:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_AA.PDF");
                            break;
                        case 6:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_REVDN.PDF");
                            break;
                        case 7:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SARAL_Revise.PDF");
                            break;
                        case 8:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SANCTION_Up_To_500_Sqmt.PDF");
                            break;
                        case 9:
                            Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised_SANCTION_Up_To_500_Sqmt.PDF");
                            break;
                    }
                    FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set REPORT_FILE_NAME = '" + Validation_FileName + "' where ID_VER = " + id + ver + ";commit;");
                    break;
                }
                System.IO.FileInfo chkfile = new System.IO.FileInfo(@"D:\From-ERP\" + dwgName);
                if (chkfile.Extension.ToUpper() != ".DWG")
                {
                    chkfile = new System.IO.FileInfo(@"D:\From-ERP\" + dwgName + ".dwg");
                }
                if (chkfile.Exists == true)
                {
                    switch (appstage)
                    {
                        case 1:
                            {
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 21 where ID_VER = " + ID_VER + ";commit;");
                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE from 1 to 21 for Application ID " + ID_VER);
                                break;
                            }
                        case 2:
                            {
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 23 where ID_VER = " + ID_VER + ";commit;");
                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE from 2 to 23 for Application ID " + ID_VER);
                                break;
                            }
                        case 3:
                            {
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 24 where ID_VER = " + ID_VER + ";commit;");
                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE from 3 to 24 for Application ID " + ID_VER);
                                break;
                            }

                    }

                }
            }
            System.Data.DataTable ProcessChkDT = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application where APP_STAGE = 47 and APP_STAGE_TEMP = 4;commit;");
            //System.Data.DataTable ProcessChkDT = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application a,System_Number s where a.ID_VER=s.ID_VER and a.APP_STAGE = 47 and a.APP_STAGE_TEMP = 4 and s.SYS_N0=7;commit;");          --------nov20
            if (ProcessChkDT.Rows.Count != 0)
            {
                string ID_VER = ProcessChkDT.Rows[0]["id_ver"].ToString();
                System.Data.DataTable DtDwg = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from drawing where ID_VER = " + ID_VER + ";");
                string timestr = DtDwg.Rows[0]["DRAWING_TIME"].ToString();
                System.Data.DataTable ExceptionRemark = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from EXCEPTIONREMARKS where EXCPTREMRKS_ID_VER = " + ID_VER + ";");
                if (ExceptionRemark.Rows.Count != 0)
                {
                    string hwnd = ExceptionRemark.Rows[0]["EXCPTREMRKS_REMARKS"].ToString();
                    DateTime dt;
                    DateTime.TryParse(timestr, out dt);
                    TimeSpan t1 = DateTime.Now.Subtract(dt);
                    if (t1.Minutes >= 5)
                    {
                        for (int ThreadNo = 0; ThreadNo < Threads.Count; ThreadNo++)
                        {
                            Thread TmpTh = Threads[ThreadNo];
                            if (TmpTh.Name == ID_VER)
                            {
                                TmpTh.Abort();
                                int processid;
                                IntPtr hnwdintptr = (IntPtr)Convert.ToInt32(hwnd);
                                int threadid = GetWindowThreadProcessId(hnwdintptr, out processid);
                                System.Diagnostics.Process Pracad = System.Diagnostics.Process.GetProcessById(processid);
                                Pracad.Kill();
                                TmpTh.Suspend();
                                log.Debug("Thread Suspended");
                                //TmpTh.Join();
                                Threads.Remove(TmpTh);
                                count--;
                                id = ID_VER.Substring(0, ID_VER.Length - 2);
                                ver = ID_VER.Substring(ID_VER.Length - 2);
                                TxtFilePath = @"d:\mcd\Report\" + id + "_" + ver + "_ValidationReport.Txt";
                                using (StreamWriter TxtFile = new StreamWriter(TxtFilePath, true))
                                {
                                    TxtFile.WriteLine("Drawing can not open, Please upload valid dwg");
                                }
                                ValidateReport vr = new ValidateReport();
                                vr.validReport(ID_VER, TxtFilePath);
                                System.IO.File.Delete(TxtFilePath);
                                DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
                                con.Open();
                                DB2Command AppstageCommand = new DB2Command("set schema " + FunctionsNvar.schema + ";select app_stage from application where id_ver =" + ID_VER + ";commit;", con);
                                Appstage = Convert.ToInt16(AppstageCommand.ExecuteScalar());
                                if ((Appstage == 23) || (Appstage == 24))
                                {
                                }
                                else
                                {
                                    FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 51,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_VER + ";commit;");
                                    //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                    AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 51 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                                    //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                }

                                DB2Command Cmd1 = new DB2Command("set schema " + FunctionsNvar.schema + ";select F_ID from application where id_ver =" + ID_VER + ";commit;", con);
                                fid = Convert.ToInt16(Cmd1.ExecuteScalar());
                                // StringBuilder Validation_FileName = new StringBuilder();
                                switch (fid)
                                {
                                    case 1:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport.PDF");
                                        break;
                                    case 2:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_CC.PDF");
                                        break;
                                    case 3:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised.PDF");
                                        break;
                                    case 4:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Regularized.PDF");
                                        break;
                                    case 5:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_AA.PDF");
                                        break;
                                    case 6:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_REVDN.PDF");
                                        break;
                                    case 7:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SARAL_Revise.PDF");
                                        break;
                                    case 8:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_SANCTION_Up_To_500_Sqmt.PDF");
                                        break;
                                    case 9:
                                        Validation_FileName.Append(id + "_" + ver + "_ValidationReport_Revised_SANCTION_Up_To_500_Sqmt.PDF");
                                        break;
                                }
                                con.Close();

                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set REPORT_FILE_NAME = '" + Validation_FileName + "' where ID_VER = " + id + ver + ";commit;");
                                break;
                            }
                        }

                    }
                }
            }
            if (count < 2)
            {
                try
                {
                    log.Debug("ChkDbTimer_Tick()- Started");
                    //  System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + "; select * from application where (APP_STAGE = 24 or APP_STAGE = 23 or APP_STAGE = 21) and APP_STAGE_TEMP IS NULL  ORDER BY ID_VER DESC ;commit;");
                    //  System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + "; select * from application where (APP_STAGE = 24 or APP_STAGE = 23 or APP_STAGE = 21) and APP_STAGE_TEMP IS NULL and APP_DATE < '2021-03-20'  ORDER BY ID_VER DESC ;commit;");
                    // System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from VM_Application_1 order by uploaded_DTM desc;commit;");

                    //  select* from application where (APP_STAGE = 24 or APP_STAGE = 23 or APP_STAGE = 21) and APP_STAGE_TEMP IS NULL  ORDER BY ID_VER DESC;
                    //                   System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + @"; SELECT b.building_plan_req_detail_id,a.app_stage, b.dwg_version, b.building_plan_req_detail_gen_id, a.app_date, fd.uploaded_dtm as updated_dtm FROM mpd2021.application a, obp_building_plan_req_detail b,obp_file_doc fd where
                    //(a.app_stage = 23 or a.app_stage = 21 or a.app_stage = 24) and a.app_stage_temp is null
                    //and a.id = b.building_plan_req_detail_id
                    //and b.building_plan_req_detail_gen_id in (select max(building_plan_req_detail_gen_id) from obp_building_plan_req_detail group by building_plan_req_detail_id)
                    //and b.building_plan_req_detail_gen_id = fd.building_plan_req_detail_gen_id
                    //and fd.file_doc_gen_id in (select max(file_doc_gen_id) from obp_file_doc where doc_type in ('Building Plan','Completion Layout','Completion DWG') group by building_plan_req_detail_gen_id)
                    //and b.status in ('1','101','150','179','240','152','61','117','151','213','221','260','107','196','9','127','178','103','161','252','425')
                    //union
                    //SELECT b.building_plan_req_detail_id,a.app_stage, b.dwg_version, b.building_plan_req_detail_gen_id, a.app_date, fd.uploaded_dtm as updated_dtm FROM mpd2021.application a, obp_building_plan_req_detail b,
                    //obp_file_doc fd where
                    //(a.app_stage = 23 or a.app_stage = 21 or a.app_stage = 24) and a.app_stage_temp is null
                    //and a.id = b.building_plan_req_detail_id
                    //and b.building_plan_req_detail_gen_id in (select max(building_plan_req_detail_gen_id) from obp_building_plan_req_detail group by building_plan_req_detail_id)
                    //and b.building_plan_req_detail_gen_id = fd.building_plan_req_detail_gen_id
                    //and fd.file_doc_gen_id in (select max(file_doc_gen_id) from obp_file_doc where doc_type in ('Building Plan','Completion Layout','Completion DWG') group by building_plan_req_detail_gen_id)
                    //and b.status in ('117','127')
                    //order by updated_dtm asc; commit;");
                    System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application where (APP_STAGE = 24 or APP_STAGE = 23 or APP_STAGE = 21) and APP_STAGE_TEMP IS NULL;commit;");
                    //System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application a, System_Number s where a.ID_VER=s.ID_VER and (a.APP_STAGE = 24 or a.APP_STAGE = 23 or a.APP_STAGE = 21) and (a.APP_STAGE_TEMP IS NULL) and s.SYS_N0=7;commit;");    

                    if (MainDtApp.Rows.Count != 0)
                    {
                        count++;
                        string ID_VER = MainDtApp.Rows[0]["id_ver"].ToString();
                        //     FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE_TEMP = 4 where ID_VER = " + ID_VER + ";commit;");
                        System.Data.DataTable DtDwg = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from drawing where ID_VER = " + ID_VER + " and dwg_ver = (select  max(dwg_ver) from drawing where ID_VER = " + ID_VER + ");");
                        if (DtDwg.Rows.Count != 0)
                        {
                            if (string.IsNullOrEmpty(DtDwg.Rows[0][6].ToString()))
                            {
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set CORNER_PLOT = '0' where ID_VER = " + ID_VER + ";commit;");
                            }
                            int app_stage = int.Parse(MainDtApp.Rows[0]["app_stage"].ToString());
                            if (app_stage == 23 || app_stage == 24)
                            {
                                //nothing now   
                            }
                            else
                            {
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 47 where ID_VER = " + ID_VER + ";commit;");
                                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 47  for Application ID " + ID_VER);
                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                            }

                            string dwgname = DtDwg.Rows[0][1].ToString();
                            if (dwgname != string.Empty)
                            {
                                string dwgPath = @"D:\From-ERP\";
                                int Ver = int.Parse(DtDwg.Rows[0][2].ToString());
                                System.IO.FileInfo newfi = new System.IO.FileInfo(dwgPath + dwgname);
                                if (newfi.Extension.ToUpper() != ".DWG")
                                {
                                    newfi = new System.IO.FileInfo(dwgPath + dwgname + ".dwg");
                                }
                                StringBuilder NewDwgPathStrBlder = new StringBuilder(newfi.FullName);
                                NewDwgPathStrBlder.Remove(NewDwgPathStrBlder.Length - 4, 4);
                                FilePathLabel.Text = "Processing drawing test --- > " + newfi.Name;
                                this.Width = FilePathLabel.Width + 34;
                                GrpBxLabl.Width = this.Width - 20;
                                this.Refresh();
                                NewDwgPathStrBlder.Append("_" + ID_VER.ToString() + ".dwg");
                                string OldDwg = newfi.FullName;
                                string NewDwg = NewDwgPathStrBlder.ToString();
                                if (System.IO.File.Exists(OldDwg) == false)
                                {
                                    switch (app_stage)
                                    {
                                        case 21:
                                            {
                                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 42,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_VER + ";commit;");
                                                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 42 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                                log.Error("ChkDbTimer_Tick()- Not obtaining the Drawing from FromERP folder for APP_Stage 1");
                                                break;
                                            }
                                        case 23:
                                            {
                                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 40,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_VER + ";commit;");
                                                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 40 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                                log.Error("ChkDbTimer_Tick()- Not obtaining the Drawing from FromERP folder for APP_Stage 2");
                                                break;
                                            }
                                        case 24:
                                            {
                                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 41,APP_STAGE_TEMP = NULL  where ID_VER = " + ID_VER + ";commit;");
                                                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 41 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                                log.Error("ChkDbTimer_Tick()- Not obtaining the Drawing from FromERP folder for APP_Stage 3");
                                                break;
                                            }

                                    }
                                    count--;
                                    return;

                                    //if (app_stage == 2 || app_stage == 3)
                                    //{
                                    //    log.Debug("ChkDbTimer_Tick()- Not obtaining the Drawing from FromERP folder for APP_Stage 2 or 3");
                                    //    count--;
                                    //    return;

                                    //}
                                    //else
                                    //{
                                    //    log.Debug("ChkDbTimer_Tick()- Updating APP_Stage to 42");
                                    //    FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 42,APP_STAGE_TEMP = NULL where ID_VER = " + ID_VER + ";commit;");
                                    //    count--;
                                    //    return;
                                    //}
                                }
                                try
                                {
                                    log.Debug("ChkDbTimer_Tick()- Coping old file to new file in FROM-ERP folder");
                                    System.IO.File.Copy(OldDwg, NewDwg, true);
                                }
                                catch (System.Exception ex)
                                {
                                    log.Error("ChkDbTimer_Tick()-Coping old file to new file in FROM-ERP folder -Error(" + ex.Message + ")");
                                    FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 1,APP_STAGE_TEMP = NULL where ID_VER = " + ID_VER + ";commit;");
                                    //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                    AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated APP_STAGE is 1 and its APP_STAGE_TEMP is NULL for Application ID " + ID_VER);
                                    //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                                    count--;
                                    return;
                                }

                                FunctionsNvar.FilePath = NewDwg;
                                opencadclass opcad = new opencadclass(NewDwg, ID_VER);
                                Thread th = new Thread(new ThreadStart(opcad.Openacad));
                                th.Name = ID_VER;
                                th.Start();
                                Threads.Add(th);
                                FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE_TEMP = 4 where ID_VER = " + ID_VER + ";commit;");
                                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                                AppStageslog.DebugLog("ChkDbTimer_Tick()-Updated  APP_STAGE_TEMP is 4 for Application ID " + ID_VER);
                                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                            }
                            else
                            {
                                count--;
                            }
                        }
                        else
                        {
                            count--;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    log.Error("ChkDbTimer_Tick()-Unable to Open Autocad -Error(" + ex.Message + ")");
                    MessageBox.Show("Error : " + ex.Message + "\n" + ex.Source);
                }
            }
            for (int ThreadNo = 0; ThreadNo < Threads.Count; ThreadNo++)
            {
                Thread TmpTh = Threads[ThreadNo];
                if (TmpTh.ThreadState == ThreadState.Stopped || TmpTh.ThreadState == ThreadState.Aborted)
                {
                    Threads.RemoveAt(ThreadNo);
                    ThreadNo--;
                    count--;
                }
            }
        }

        private void ChkDbDocTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                log.Debug("ChkDbDocTimer_Tick()- Started");
                //DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
                //con.Open();
                //DB2Command cmd = new DB2Command("set schema " + FunctionsNvar.schema  + ";select * from application where APP_STAGE = 6;commit;", con);
                //DB2DataReader reader = cmd.ExecuteReader();
                System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application where APP_STAGE = 68;commit;"); //Selecting records with app_stage 68 changed by Kiran  on 21st Aug 2013.
                //System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application a,System_Number s  where a.ID_VER=s.ID_VER and a.APP_STAGE = 66 and s.SYS_N0=7;commit;");

                /*This condition was commented by Kiran Bishaj on 6th Sept 2013.
                if (MainDtApp.Rows.Count == 0)
                {
                    MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application where APP_STAGE = 67;commit;");
                    //MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application a, System_Number s where a.ID_VER=s.ID_VER and a.APP_STAGE = 67 and SYS_N0=7;commit;");
                
                 }
                */

                if (MainDtApp.Rows.Count != 0)
                {

                    string ID = MainDtApp.Rows[0]["id_ver"].ToString();
                    string appid = ID.Substring(0, ID.Length - 2);
                    string Ver = ID.Substring(ID.Length - 2);
                    System.Data.DataTable DtDwg = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from drawing where ID_VER = " + ID + "; commit;");
                    if (DtDwg.Rows.Count != 0)
                    {
                        FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 48,APP_STAGE_TEMP = NULL where ID_VER = " + ID + ";commit;");
                        //-->>> Included Appstage log on 29-05-2013 By Kiran  
                        AppStageslog.DebugLog("ChkDbDocTimer_Tick()-Updated APP_STAGE is 48 and its APP_STAGE_TEMP is NULL for Application ID " + ID);
                        //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                        string dwgname = DtDwg.Rows[0][1].ToString();
                        string dwgPath = @"d:\mcd\Report\";//DtDwg.Rows[0][4].ToString();

                        System.Data.DataTable DtBuildType = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select  BLDG_TYPE_ID from application where ID_VER = " + ID + "; commit;");
                        int buildTypeId = int.Parse(DtBuildType.Rows[0][0].ToString());
                        FilePathLabel.Text = "Processing Report for --- > " + dwgname;
                        this.Width = FilePathLabel.Width + 34;
                        GrpBxLabl.Width = this.Width - 20;
                        this.Refresh();
                        //System.Windows.Forms.Application.DoEvents();
                        //object filename = dwgPath +  ID  ;
                        ReportDoc rdoc = new ReportDoc();
                        bool approved = rdoc.report(dwgPath, ID.ToString(), buildTypeId);

                        //cmd2 = new DB2Command("set schema " + FunctionsNvar.schema  + ";update application set APP_STAGE = 5 where ID = " + ID + ";commit;", con);
                        //cmd2.ExecuteNonQuery();
                        if (approved == true)
                        {
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 52,APP_STAGE_TEMP = NULL where ID_VER = " + ID + ";commit;");
                            //-->>> Included Appstage log on 29-05-2013 By Kiran  
                            AppStageslog.DebugLog("ChkDbDocTimer_Tick()-Updated APP_STAGE is 52 and its APP_STAGE_TEMP is NULL for Application ID " + ID);

                            //<<<-- Included Appstage log on 29-05-2013 By Kiran 
                            DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
                            con.Open();
                            DB2Command Cmd1 = new DB2Command("set schema " + FunctionsNvar.schema + ";select F_ID from application where id_ver =" + ID + ";commit;", con);
                            int fid = Convert.ToInt16(Cmd1.ExecuteScalar());
                            StringBuilder ByeLaw_FileName = new StringBuilder();
                            switch (fid)
                            {
                                case 1:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport.PDF");
                                    break;
                                case 2:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_CC.PDF");
                                    break;
                                case 3:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Revised.PDF");
                                    break;
                                case 4:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Regularized.PDF");
                                    break;
                                case 5:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_AA.PDF");
                                    break;
                                case 6:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_REVDN.PDF");
                                    break;
                                case 7:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_SARAL_Revise.PDF");
                                    break;
                                case 8:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_SANCTION_Up_To_500_Sqmt.PDF");
                                    break;
                                case 9:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Revised_SANCTION_Up_To_500_Sqmt.PDF");
                                    break;
                            }
                            con.Close();

                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set REPORT_FILE_NAME = '" + ByeLaw_FileName + "' where ID_VER = " + ID + ";commit;");

                            log.Debug("In-order Bye-Law report generated successfully for drwaing:- (" + dwgname + ") with id : " + ID + " ");
                        }
                        else
                        {
                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set APP_STAGE = 53,APP_STAGE_TEMP = NULL where ID_VER = " + ID + ";commit;");
                            //-->>> Included Appstage log on 29-05-2013 By Kiran  
                            AppStageslog.DebugLog("ChkDbDocTimer_Tick()-Updated APP_STAGE is 53 and its APP_STAGE_TEMP is NULL for Application ID " + ID);
                            //<<<-- Included Appstage log on 29-05-2013 By Kiran 

                            DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
                            con.Open();
                            DB2Command Cmd1 = new DB2Command("set schema " + FunctionsNvar.schema + ";select F_ID from application where id_ver =" + ID + ";commit;", con);
                            int fid = Convert.ToInt16(Cmd1.ExecuteScalar());
                            StringBuilder ByeLaw_FileName = new StringBuilder();
                            switch (fid)
                            {
                                case 1:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport.PDF," + appid + "_" + Ver + "_Error.dwg");
                                    break;
                                case 2:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_CC.PDF," + appid + "_" + Ver + "_Error_CC.dwg");
                                    break;
                                case 3:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Revised.PDF," + appid + "_" + Ver + "_Error_Revised.dwg");
                                    break;
                                case 4:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Regularized.PDF," + appid + "_" + Ver + "_Error_Regularized.dwg");
                                    break;
                                case 5:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_AA.PDF," + appid + "_" + Ver + "_Error_AA.dwg");
                                    break;
                                case 6:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_REVDN.PDF," + appid + "_" + Ver + "_Error_REVDN.dwg");
                                    break;
                                case 7:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_SARAL_Revise.PDF," + appid + "_" + Ver + "_Error_SARAL_Revise.dwg");
                                    break;
                                case 8:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_SANCTION_Up_To_500_Sqmt.PDF," + appid + "_" + Ver + "_Error_SANCTION_Up_To_500_Sqmt.dwg");
                                    break;
                                case 9:
                                    ByeLaw_FileName.Append(appid + "_" + Ver + "_ByeLawReport_Revised_SANCTION_Up_To_500_Sqmt.PDF," + appid + "_" + Ver + "_Error_Revised_SANCTION_Up_To_500_Sqmt.dwg");
                                    break;
                            }
                            con.Close();

                            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update DRAWING set REPORT_FILE_NAME = '" + ByeLaw_FileName + "' where ID_VER = " + ID + ";commit;");



                            log.Debug("Not In-order Bye-Law report generated successfully for drwaing:- (" + dwgname + ") with id : " + ID + " ");
                        }
                        FilePathLabel.Text = "Process complete for Report --- > " + dwgname;
                        log.Debug("Bye-Law report generated successfully for drwaing:- (" + dwgname + ") with id : " + ID + " ");
                        this.Width = FilePathLabel.Width + 34;
                        GrpBxLabl.Width = this.Width - 20;
                    }

                }
                //con.Close();
            }
            catch (System.Exception ex)
            {
                log.Error("ChkDbDocTimer_Tick()-Unable to Open Autocad -Error(" + ex.Message + ")");

                MessageBox.Show("Error : " + ex.Message + "\n" + ex.StackTrace + "\n" + ex.Source);
            }
        }

        private void StopBttn_Click(object sender, EventArgs e)
        {
            log.Debug("Stop Button clicked");
            DialogResult dr = MessageBox.Show("Do you want to stop execution?", "Mcd Building Plan", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (dr.Equals(DialogResult.Yes) == true)
            {
                count = 100;
                for (int ThNo = 0; ThNo < Threads.Count; ThNo++)
                {
                    Thread Th = Threads[ThNo];
                    Th.Suspend();
                    log.Debug("Thread Suspended");
                }
            }

        }

        private void GrpBxLabl_Enter(object sender, EventArgs e)
        {

        }
        public int AppnDbquery(string ID_VER)
        {
            int intIdAppStage;
            DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
            con.Open();
            DB2Command AppstageCommand = new DB2Command("set schema " + FunctionsNvar.schema + ";select app_stage from application where id_ver =" + ID_VER + ";commit;", con);
            intIdAppStage = Convert.ToInt16(AppstageCommand.ExecuteScalar());
            con.Close();
            return intIdAppStage;
        }
        /// <summary>
        /// The following Modifications were done in startupRecovery()
        /// 1.Included Switch case to update ids in Appstage 23 and 40,24 and 41 to appstage 2,3 respectively and remaining to appsatage 1. 
        /// 2.Introduced a 2d array with all table,column names where data to be deleted for id's with app_stage 1.
        /// 3.Created function deleteData(string table, string idColumn, string idver) to replace deleteData(string id).
        /// By Kiran Bishaj on 31st July 2013.
        /// </summary>
        private void startupRecovery()
        {
            log.Debug("startupRecovery()- Started");

            FilePathLabel.Text = "Recovery started test.";
            this.Refresh();
            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set app_stage = 1,APP_STAGE_TEMP=null " +
                                                "where APP_STAGE = 42 or APP_STAGE = 43 or APP_STAGE = 44 or APP_STAGE = 21;commit;");
            //-->>> Included Appstage log on 29-05-2013 By Kiran  
            AppStageslog.DebugLog("startupRecovery()- Updated app_stage to 1 and App_Stage_temp to NULL ,If App_Stage is 42 or 43 or 44 or 21");
            //<<<-- Included Appstage log on 29-05-2013 By Kiran 
            System.Data.DataTable FileDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select * from application " +
                                                "where APP_STAGE = 22 or APP_STAGE = 23 or APP_STAGE = 24  or  APP_STAGE = 40 or  APP_STAGE = 41 or  APP_STAGE = 45 or " +
                                                "APP_STAGE = 47 or APP_STAGE = 48 or APP_STAGE = 66 or  APP_STAGE = 67;commit;");
            log.Debug("startupRecovery()- Selected the records having app_stage=22 or 45 or 47 or 66");
            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";delete from  EXCEPTIONREMARKS;commit;");
            log.Debug("startupRecovery()- Deleted data from Exceptionremarks Table");
            for (int i = 0; i < FileDtApp.Rows.Count; i++)
            {
                string ID_ver = FileDtApp.Rows[i]["id_ver"].ToString();

                string[,] tablecolumn = new string[,]
    {

            {"DA_BASEMENT","BASE_ID_VER"},
            {"DA_BATH_WATERCLOSET_ROOM","BATH_ID_VER"},
            {"DA_CORRIDORS","ID_VER"},
            {"DA_FIREESCAPE_STAIRCASE","ID_VER"},
            {"DA_GSQ_GARAGE","GARAGE_ID_VER"},
            {"DA_LIFTLOBBY","ID_VER"},
            {"DA_LIFTPIT","LP_ID_VER"},
            {"DA_LOFT","ID_VER"},
            {"DA_MEZZANINE","MEZ_ID_VER"},
          {"DA_NOTIFIED_COMMERCIAL_AREA", "NOTIFIED_COMMAREA_ID_VER"},
          {"DA_NOTIFIED_RAMPS", "RAMP_ID_VER"},
          {"DA_NOTIFIED_STAIRCASE", "NS_ID_VER"},
            {"DA_PARKING","ID_VER"},
            {"DA_PASSAGEWAYS_WT","ID_VER"},
            {"DA_PERGOLA","ID_VER"},	  
		  //{"DA_PERMITFEE_AREA", "FEES_FLR_ID_VER"},--Down
            {"DA_RES_BALCONY","BALCONY_ID_VER"},
            {"DA_RES_BNDRY_WALL","RBW_ID_VER"},
            {"DA_RES_BUILDING","RESPLTH_ID_VER"},
            {"DA_RES_CANOPY","CANOPY_ID_VER"},
          {"DA_RES_COMMERCIAL_FEATURES", "CF_ID_VER"},
          {"DA_RES_COMMERCIAL_SUBFEATURES", "CSF_ID_VER"},
          //{"DA_RES_COV_FEE", "ID_VER"},	 --Down
            {"DA_RES_CUPBOARD_SHELVES","ID_VER"},
            {"DA_RES_DOOR_WINDOW","RESDRW_ID_VER"},
            {"DA_PERMITFEE_AREA","FEES_FLR_ID_VER"},
		  //{"DA_RES_DWELLING", "RESDU_ID_VER"}, --Down
          //{"DA_RES_FLOOR", "RESFLR_ID_VER"},	  --Down
          //{"DA_RES_FLOOR_HT", "RESFLRHT_ID_VER"}, --Down
            {"DA_RES_GARAGE","RG_ID_VER"},
            {"DA_RES_HABITABLE_ROOM","RESHABR_ID_VER"},
            {"DA_RES_HAND_RAILS","RHR_ID_VER"},
            {"DA_RES_HEADROOM_STAIRCASE","RHS_ID_VER"},
            {"DA_RES_INTERIOR_COURTYARD","RIC_ID_VER"},
		  //{"DA_RES_INTERMEDIATE_FEE", "ID_VER"}, --Down
            {"DA_RES_LEDGE_TAND","ID_VER"},

		  //{"DA_RES_OPENAREA", "OPENAREA_ID_VER"},--Down
            {"DA_RES_PANTRIES","RP_ID_VER"},	
		  //{"DA_RES_PLOT", "RESPLT_ID_VER"},--Down
            {"DA_RES_PPT_WALL","RPW_ID_VER"},
            {"DA_RES_PRV_LIFT","RPL_ID_VER"},
          {"DA_RES_RGH_COMMUNITYHALLS", "ID_VER"},
          {"DA_RES_RGH_EWS", "ID_VER"},
          {"DA_RES_RGH_EWSDWELLING", "ID_VER"},
            {"DA_RES_ROOM_D_W","RESRDW_ID_VER"},
            {"DA_RES_SETBACK","RESSBID_VAR"},
            {"DA_RES_SPIRAL_STAIRS","RSS_ID_VER"},
            {"DA_RES_SQ_BLOCK","RESSQ_ID_VER"},
            {"DA_RES_STAIRWAYS","RS_ID_VER"},
            {"DA_RES_WEATHER_SHADE","RWS_ID_VER"},
            {"DA_SERVANT_QUARTERS","ID_VER"},
            {"DA_STILT","ST_ID_VER"},
            {"DA_STORE_ROOM","SR_ID_VER"},
            {"DA_VENT_SHAFT","VSHAFT_ID_VER"}, 
		  //{"DA_VERANDA", "ID_VER"},	--Down
          //{"ERROR_SUMMARY", "ID_VER"},--Down
          {"EXCEPTIONREMARKS", "EXCPTREMRKS_ID_VER"},
            {"GENERAL_ERRORS","ID_VER"},	
		  //{"I117_ERROR_SUMMARY", "ID_VER"},-Down
            {"I117_RE_FEE","ID_VER"},
            {"I117_RE_SETBACK","ID_VER"},
          {"RE102_PRORATA", "ID_VER"},
          //{"RES_INTERMEDIATE_FLOOR_HT", "INTRMDT_FLRHT_ID_VER"},--Down
            {"RE_BALCONY","ID_VER"},
            {"RE_CANOPY","ID_VER"},
            {"RE_CANOPY_TOTAL","ID_VER"},
          {"RE_CARLIFT", "ID_VER"},
          {"RE_COMMERCIAL_FEATURES_COUNT", "ID_VER"},
            {"RE_CORRIDORS","ID_VER"},
            {"RE_COURTYARD","ID_VER"},
            {"RE_COVERAGE","ID_VER"},
             {"RE_COVERAGE_DIFF", "ID_VER"},
            {"RE_DWELLING_UNIT_COUNT","ID_VER"},
            {"RE_FAR","ID_VER"},
          {"RE_FEES_DIFFERENCE", "ID_VER"},
            {"RE_FIREESCAPE_STAIRCASE","ID_VER"}, 
		  //{"RE_FLOOR_WISE_PERMIT_FEE", "ID_VER"},--undefined in Production
            {"RE_HEIGHT","ID_VER"},
            {"RE_INDIVIDUAL_DWELLING_COUNT","ID_VER"},
            {"RE_LOFT","ID_VER"},
            {"RE_LOFT_HT","ID_VER"},
            {"RE_NOTE","ID_VER"},
          {"RE_NOTIFIED_DWELLING_UNIT_COUNT", "ID_VER"},
		  //{"RE_NOTIFIED_ERROR_SUMMARY", "ID_VER"},--Down
		  {"RE_NOTIFIED_RAMPS", "ID_VER"},
          {"RE_NOTIFIED_STAIRCASE", "ID_VER"},
          {"RE_OFFICE", "ID_VER"},
            {"RE_PARKING","ID_VER"},
            {"RE_PARKING_TOTAL_NO","ID_VER"},
            {"RE_PASSAGEWAYS_WT","ID_VER"},
            {"RE_PERGOLA","ID_VER"},
            {"RE_PERGOLA_TOTAL","ID_VER"}, 
		  //{"RE_RES_BASEMENT","ID_VER"},--Down
            {"RE_RES_BNDRY_WALL","ID_VER"},
            {"RE_RES_CUPBOARD_SHELVES","ID_VER"},
            {"RE_RES_FEE","ID_VER"},
		   // {"RE_RES_GARAGE","ID_VER"},--Down
            {"RE_RES_HAND_RAILS","ID_VER"},
            {"RE_RES_HEADROOM_STAIRCASE","ID_VER"},
            {"RE_RES_LEDGE_TAND","ID_VER"},
            {"RE_RES_LEDGE_TAND_HT","ID_VER"},
          {"RE_RES_NOTIFIED_FEES","ID_VER"},
            {"RE_RES_PANTRIES","ID_VER"},
            {"RE_RES_PARAPET_WALL","ID_VER"},
            {"RE_RES_PROVSION_LIFT","ID_VER"},
          {"RE_RES_RGH_COMMUNITYHALLS","ID_VER"},
          {"RE_RES_RGH_EWSDWELLING","ID_VER"},
            {"RE_RES_SPIRAL_STAIRS","ID_VER"},
            {"RE_RES_STAIRWAYS","ID_VER"}, 
		  //{"RE_RES_STILT","ID_VER"},	 --Down
            {"RE_RES_STORE_ROOM","ID_VER"},	
		  //{"RE_RES_TOTAL_CUPBOARD_SHELVES","ID_VER"}, --Down
            {"RE_RES_WEATHER_SHD","ID_VER"},
            {"RE_ROOMS","ID_VER"},
            {"RE_SERVANT_QUARTERS","ID_VER"},
            {"RE_SETBACK","ID_VER"},
            {"RE_SHAFT","ID_VER"},	
		  //{"RE_SHOP","ID_VER"},
            {"RE_VENTILATION","ID_VER"},	
		  //{"RE_VERANDA","ID_VER"}, --Down
            {"ERROR_SUMMARY","ID_VER"},
            {"I117_ERROR_SUMMARY","ID_VER"},
            {"RE_NOTIFIED_ERROR_SUMMARY","ID_VER"},
            {"EXCEPTIONREMARKS","EXCPTREMRKS_ID_VER"},
            {"DA_RES_DWELLING","RESDU_ID_VER"},
            {"DA_RES_FLOOR","RESFLR_ID_VER"},
            {"DA_RES_FLOOR_HT","RESFLRHT_ID_VER"},
            {"RE_RES_STILT","ID_VER"},
            {"RE_RES_GARAGE","ID_VER"},
            {"RE_RES_TOTAL_CUPBOARD_SHELVES","ID_VER"},
            {"DA_RES_INTERMEDIATE_FEE","ID_VER"},
            {"DA_RES_COV_FEE","ID_VER"},
            {"DA_RES_OPENAREA","OPENAREA_ID_VER"},
            {"DA_VERANDA","ID_VER"},
            {"RE_VERANDA","ID_VER"},
            {"RES_INTERMEDIATE_FLOOR_HT","INTRMDT_FLRHT_ID_VER"},
            {"RE_RES_BASEMENT","ID_VER"},
            {"DA_RES_PLOT","RESPLT_ID_VER"},

        };




                Int16 appstage = (Int16)FileDtApp.Rows[i]["APP_STAGE"];
                switch (appstage)
                {
                    case 23:

                    case 40:
                        FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set app_stage = 2,APP_STAGE_TEMP=null " +
                                                 "where id_ver = " + ID_ver + " ;commit;");
                        break;
                    case 24:
                    case 41:
                        FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set app_stage = 3,APP_STAGE_TEMP=null " +
                                                "where id_ver = " + ID_ver + " ;commit;");
                        break;
                    default:

                        for (int k = 0; k <= tablecolumn.GetUpperBound(0); k++)
                        {
                            string s1 = tablecolumn[k, 0]; // Table names
                            string s2 = tablecolumn[k, 1]; //id Column names

                            deleteData(s1, s2, ID_ver);
                        }
                        FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";update application set app_stage = 1,APP_STAGE_TEMP=null " +
                                                 "where id_ver = " + ID_ver + " ;commit;");
                        break;

                }

                System.Data.DataTable MainDtApp = FunctionsNvar.Executequery("set schema " + FunctionsNvar.schema + ";select APP_STAGE from application where ID_VER = " + ID_ver + ";commit;");
                int updatedAppstage = int.Parse(MainDtApp.Rows[0][0].ToString());
                //-->>> Included Appstage log on 29-05-2013 By Kiran  
                AppStageslog.DebugLog("startupRecovery()- Updated APP_STAGE to " + updatedAppstage + "  and APP_STAGE_TEMP to NULL for ID_VER -(" + ID_ver + ")");
                //<<<-- Included Appstage log on 29-05-2013 By Kiran 
            }
            FilePathLabel.Text = "Recovery Completed";
            log.Debug("startupRecovery()- Completed");
        }

        private void deleteData(string table, string idColumn, string idver)
        {
            FunctionsNvar.ExecuteNquery("set schema " + FunctionsNvar.schema + ";delete from " + table + " where " + idColumn + " = + " + idver + " ;commit;");
        }



        private void MainFrm_Shown(object sender, EventArgs e)
        {
            log.Debug("MainFrm_Shown()- Started");
            startupRecovery();
            ChkDbAcadTimer.Enabled = true;
            ChkDbDocTimer.Enabled = true;
            log.Debug("MainFrm_Shown()- Completed");
        }

        private void MainFrm_Load(object sender, EventArgs e)
        {

        }

    }

}
