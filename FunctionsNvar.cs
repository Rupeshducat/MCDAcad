using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using log4net;
using System.Configuration;

namespace MCD
{
    public static  class FunctionsNvar
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static string Constr = "Server=192.168.11.140:50000; Database=MCDPROD; UID=mcdacusr; PWD=Admin;";
        public static string schema = "MPD2021";
        
        public static string FilePath;
        public static string AppId;
        public static string DbStatus;
        public  const int _messageID = -1163005939;
        public  const int  WM_COPYDATA = 0x4A;

        internal static bool ExecuteNquery(string exestr)
        {
            try
            {
                log.Debug("ExecuteNquery() - Started");
                DB2Connection con = new DB2Connection(Constr);
                con.Open();
                DB2Command cmd = new DB2Command(exestr, con);
                try
                {
                    log.Debug("ExecuteNquery() - Started");
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log.Error("ExecuteNquery()-Error Occured in Database Connection-Error(" + ex.Message + ")");
                    return false;
                }
                finally
                {
                   
                    con.Close();
                    con.Dispose();
                }
                log.Debug("ExecuteNquery() - Ended");
            }
            catch (System.Exception ex)
            {
                log.Error("ExecuteNquery()-Error Occured in Database Connection-Error(" + ex.Message + ")");
                System.Windows.Forms.MessageBox.Show("Error" + "\n" + ex.Message + "\n" + ex.StackTrace);
             
                System.Windows.Forms.Application.Exit();

            }
            return true;
        }

        internal static DataTable Executequery(string exestr)
        {
            DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
            DataTable dt = new DataTable();
            try
            {
                log.Debug("ExecuteNquery() - Started");
                con.Open();
                //DB2Command cmd = new DB2Command(exestr,con );
                //DB2DataReader  reader;
                //try
                //{
                //    reader = cmd.ExecuteReader();
                //}
                //catch
                //{
              //  con.Close();
                //con.Dispose();
                //    return null;
                //}
                //DB2DataAdapter da = new DB2DataAdapter(exestr, FunctionsNvar.Constr);
                DB2DataAdapter da = new DB2DataAdapter(exestr, con);
                DataSet ds = new DataSet();
                try
                {
                    
                    da.Fill(ds, "MCD");
                    dt = ds.Tables["MCD"];
                }
                finally
                {
                    con.Dispose();
                }
            }
            catch (System.Exception ex)
            {
                log.Error("ExecuteNquery()-Error Occured in Database Connection-Error(" + ex.Message + ")");
                MessageBox.Show("Error Occured while connecting to database. Please make sure the database connecting and start the program. \n" + ex.Message, "Db not connecting", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //string msg = "<p class=MsoNormal align=center style='text-align:center'><b><span style='font-size:20.0pt;color:#C00000'>" +
                //               "Error Occured while connecting to database. Please make sure the database connecting and start the program."+
                //               "</span><span style='color:#C00000'><o:p></o:p></span></b></p>" +
                //               "<p class=MsoNormal><b><span style='color:#1F497D'>" + ex.Message + "</span></b></p>" +
                //               "<p class=MsoNormal><b><span style='color:#1F497D'>" + ex.StackTrace + "</span></b></p>";
                //MailToMCD.sendMail(msg);
                Environment.Exit(0);   
            }
            return dt;
        }
    
    }
}
