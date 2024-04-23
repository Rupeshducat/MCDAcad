/// <summary>
/// Included AppStageslog to MCDAcad on 29-05-2013 By Kiran to track appstage of Application ID's
/// </summary>
        
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;




namespace MCD
{
    public class AppStageslog
    {
        
        public static void DebugLog(string text)
        {
            try
            {
                string DebugTxtFilePath = @"D:\Logs\AppStages.txt";

                writetext(DebugTxtFilePath, text);
            }
           catch (System.Exception ex)
            {
                ErrorLog(ex.Message + ex.StackTrace);
            }

        }
        public static void ErrorLog(string text)
        {
            try
            {
                string ErrorTxtFilePath = @"D:\Logs\AppStages-Error.txt";
                writetext(ErrorTxtFilePath, text);
            }
            catch (System.Exception ex)
            {
                ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        public static void writetext(string path, string text)
        {
    
            string directory = @"D:\Logs";
            int i = 0;
            try
            {
                FileInfo f = new FileInfo(path);
               
                
                if (f.Exists == true)
                {                    
                    long s1 = f.Length;

                    if (s1 > 2097152)
                    {
                        for (i = 1; i < 100; i++)
                        {
                            string filewthoutex = System.IO.Path.GetFileNameWithoutExtension(path);
                            string newpath = @"D:\Logs\" + filewthoutex + i + ".txt";
                            FileInfo f1 = new FileInfo(newpath);

                            if (f1.Exists == false)
                            {  
                                try
                                {
                               
                                 f.MoveTo(newpath);
                                
                                    writetext(path, text);
                                }
                                catch (System.Exception ex)
                                {
                                    ErrorLog(ex.Message + ex.StackTrace);
                                }
                                break;

                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            using (FileStream file = new FileStream(path,FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                            {
                                using (StreamWriter TxtFile = new StreamWriter(file))
                                {
                                    
                                    TxtFile.WriteLine(DateTime.Now.ToString() + " " + text);
                                    //writetext(path, text);  
                                    //sw.WriteLine(text);
                                    //sw.Close();
                                }
                                //file.Flush();

                            }
                           

                        }
                        finally
                        {
                        }
                    }
                }
                else
                {
                    using (FileStream file = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                    {
                        using (StreamWriter sw = new StreamWriter(file))
                        {
                            //writetext(path, text);  
                            //sw.WriteLine(text);
                            sw.Close();
                            writetext(path, text);
                        }
                        //file.Close();
                        //file.Flush();
                    }
                }
                
                
            }
            catch (System.IO.IsolatedStorage.IsolatedStorageException)
            {
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                System.IO.Directory.CreateDirectory(directory);
    
                writetext(path, text);

            }
            catch (System.IO.FileNotFoundException)
            {
                System.IO.File.Create(path);

                writetext(path, text);
                //throw;
            }
            catch (System.IO.InternalBufferOverflowException)
            {
               
            }
            catch (System.IO.PathTooLongException)
            {

            }
        }
       
        

    }
}
