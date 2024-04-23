using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using wd = Microsoft.Office.Interop.Word;
using IBM.Data.DB2;
using System.Data.SqlTypes;

namespace MCD
{
    class ValidateReport
    {
        public void validReport(string APP_ID, string Filename)
        {
            bool retval = false;
            DB2Connection con = new DB2Connection(FunctionsNvar.Constr);
            try
            {
                con.Open();
            }
            catch (Exception ex)
            {

                System.Windows.Forms.MessageBox.Show("Server Connection Not found please contact administrator \n error: " + ex.StackTrace, "MCD Building Plan",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
            Application WordApp = new Application();
            WordApp.Visible = false;
            object readOnly = false;
            object isVisible = true;
            object missing = System.Reflection.Missing.Value;
            Document doc = WordApp.Documents.Add(ref missing, ref missing, ref missing, ref isVisible);
            object savechanges = false;
            doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            doc.Sections[1].Borders.Enable = 1;
            Object start = Type.Missing;
            Object end = Type.Missing;
            Object unit = Type.Missing;
            Object count = Type.Missing;
            doc.Range(ref start, ref end).
            Delete(ref unit, ref count);
            start = 0;
            end = 0;
            object oEndOfDoc = "\\endofdoc";

            string imagePath = @"D:\MCD\src\MCD-LOGO.bmp";

            Range rng = doc.Range(ref start, ref end);
            rng.InsertParagraphAfter();
            rng.InlineShapes.AddPicture(imagePath, ref missing, ref missing, ref missing);
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            imagePath = @"D:\MCD\src\mcd-hindi.bmp";
            //rng.InsertParagraphBefore();
            rng.InsertParagraphAfter();
            rng.InlineShapes.AddPicture(imagePath, ref missing, ref missing, ref missing);

            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //rng.InsertParagraphBefore(); 
            //rng.InsertParagraphAfter();

            rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //rng.InsertParagraphBefore();
            rng.InsertParagraphAfter();
            rng.Paragraphs.Add(ref missing);
            //rng.InsertParagraphAfter();
            rng.Text = "Municipal Corporation of Delhi";

            rng.Font.Name = "Verdana";
            rng.Font.Size = 16;
            rng.Font.Color = WdColor.wdColorAqua;
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //rng.Font.Position = 1;

            //rng.InsertParagraphBefore ();
            rng.InsertParagraphAfter();
            rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //rng.InlineShapes.AddHorizontalLineStandard(ref missing);

            object orng = rng;
            InlineShape horizontalLine = doc.InlineShapes.AddHorizontalLineStandard(ref orng);
            horizontalLine.Width = 400;
            rng.Font.Color = WdColor.wdColorAqua;
            rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            //rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            DB2Command AppCmd = new DB2Command("set schema " + FunctionsNvar.schema + ";SELECT * FROM application where ID_VER = " + APP_ID + ";commit;", con);
            DB2DataReader Appreader = AppCmd.ExecuteReader();
            int buildTypeId = 0;
            bool appRead = false;
            while (Appreader.Read())
            {
                buildTypeId = int.Parse(Appreader.GetValue(4).ToString());
                appRead = true;
            }
            if (appRead == false)
            {
                WordApp.Quit(ref savechanges, ref  missing, ref missing);
                return;
            }
            string ID = APP_ID;
            ID = ID.Remove(ID.Length - 2);
            DB2Command PropCmd = new DB2Command("set schema " + FunctionsNvar.schema + ";SELECT * FROM PROPERTY_DETAILS where ID = " + ID + ";commit;", con);
            DB2DataReader Propreader = PropCmd.ExecuteReader();
            if (Propreader.Read() == false)
            {
                WordApp.Quit(ref savechanges, ref  missing, ref missing);
                return;

            }
            DB2Command DwgCmd = new DB2Command("set schema " + FunctionsNvar.schema + ";SELECT * FROM DRAWING where ID = " + ID + "  order by  DWG_VER DESC;commit;", con);
            DB2DataReader Dwgreader = DwgCmd.ExecuteReader();
            if (Dwgreader.Read() == false)
            {
                WordApp.Quit(ref savechanges, ref  missing, ref missing);
                return;

            }
            DB2Command Cmd1 = new DB2Command("set schema " + FunctionsNvar.schema + ";select F_ID from application where id_ver =" + APP_ID + ";commit;", con);
            int fid = Convert.ToInt16(Cmd1.ExecuteScalar());
            object objAutoFitFixed2 = WdAutoFitBehavior.wdAutoFitWindow;
            Table tbl2 = doc.Tables.Add(rng, 4, 4, ref missing, ref objAutoFitFixed2);
            tbl2.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto;
            tbl2.Range.Font.Size = 8;
            Object style = "Table Grid 1";
            tbl2.set_Style(ref style);

            //-->>Included Two New columns Architect Name and Architect CA No in the Report on 24th Sept 2013 By Kiran Bishaj.
            tbl2.Cell(1, 1).Range.Text = "Architect Name :";
            tbl2.Cell(1, 1).Range.Bold = 1;
            tbl2.Cell(1, 2).Range.Text = Propreader.GetValue(3).ToString();
            tbl2.Cell(1, 3).Range.Text = "Architect CA No :";
            tbl2.Cell(1, 3).Range.Bold = 1;
            tbl2.Cell(1, 4).Range.Text = Propreader.GetValue(4).ToString();

            tbl2.Cell(2, 1).Range.Text = "Applicant Name :";
            tbl2.Cell(2, 1).Range.Bold = 1;
            tbl2.Cell(2, 2).Range.Text = Propreader.GetValue(2).ToString();
            tbl2.Cell(2, 3).Range.Text = "Drawing Name :";
            tbl2.Cell(2, 3).Range.Bold = 1;
            tbl2.Cell(2, 4).Range.Text = Dwgreader.GetValue(1).ToString();
            tbl2.Cell(3, 1).Range.Text = "Building Type :";
            tbl2.Cell(3, 1).Range.Bold = 1;

            if (buildTypeId == 101)
            {
                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential_SARAL_Revise ";
                }
            }
            if (buildTypeId == 102)
            {

                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "URC_SARAL_Revise ";
                }
            }
            if (buildTypeId == 103)
            {

                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Village Abadi_SARAL_Revise ";
                }
            }
            if (buildTypeId == 104)
            {

                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "City Area _SARAL_Revise ";
                }
            }
            if (buildTypeId == 105)
            {

                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Standard Plan_SARAL_Revise ";
                }
            }
            if (buildTypeId == 106)
            {

                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential Group Housing ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential Group Housing_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential Group Housing_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Residential Group Housing_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Resedential Group Housing_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Resedential Group Housing_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Resedential Group Housing_SARAL_Revise ";
                }
            }

            if (buildTypeId == 107)
            {
                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified-Commercial_SARAL_Revise ";
                }
            }

            if (buildTypeId == 108)
            {
                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Notified - MLU_SARAL_Revise ";
                }
            }


            else if (buildTypeId == 117)
            {
                if (fid == 1)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial ";
                }
                if (fid == 2)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_CC ";
                }
                if (fid == 3)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_Revised ";
                }
                if (fid == 4)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_Regularized ";
                }
                if (fid == 5)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_AA ";
                }
                if (fid == 6)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_REVDN ";
                }
                if (fid == 7)
                {
                    tbl2.Cell(3, 2).Range.Text = "Industrial_SARAL_Revise";
                }
            }
            tbl2.Cell(3, 3).Range.Text = "Plot Area :";
            tbl2.Cell(3, 3).Range.Bold = 1;
            tbl2.Cell(3, 4).Range.Text = Dwgreader.GetDecimal(13).ToString();
            tbl2.Cell(4, 1).Range.Text = "Application ID :";
            tbl2.Cell(4, 1).Range.Bold = 1;
            tbl2.Cell(4, 2).Range.Text = ID;
            tbl2.Cell(4, 3).Range.Text = "Date :";
            tbl2.Cell(4, 3).Range.Bold = 1;
            tbl2.Cell(4, 4).Range.Text = Dwgreader.GetDate(4).ToShortDateString();
            tbl2.Cell(5, 1).Range.Text = "Address :";
            tbl2.Cell(5, 1).Range.Bold = 1;
            tbl2.Cell(5, 2).Range.Text = Propreader.GetValue(1).ToString();
            //rng.InsertParagraphAfter();
            tbl2.Cell(5, 3).Range.Text = " In order/Not in order :";
            tbl2.Cell(5, 3).Range.Bold = 1;
            tbl2.Cell(5, 4).Range.Text = "Not in order";
            //<<--Included Two New columns Architect Name and Architect CA No in the Report on 24th Sept 2013 By Kiran Bishaj.

            /*******************************for summary Table******************/

            Paragraph oPara4;
            rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            object parang = rng;
            oPara4 = doc.Content.Paragraphs.Add(ref parang);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "Validation errors";
            oPara4.Range.Font.Underline = WdUnderline.wdUnderlineSingle;
            oPara4.Range.Font.Size = 16;
            oPara4.Range.Font.Color = WdColor.wdColorDarkRed;
            oPara4.Range.InsertParagraphAfter();

            System.IO.StreamReader fname = new System.IO.StreamReader(Filename);
            int no = 0;
            while (fname.EndOfStream == false)
            {
                string error = fname.ReadLine();
                no++;
                rng = doc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                parang = rng;
                oPara4 = doc.Content.Paragraphs.Add(ref parang);
                //oPara4.Range.InsertParagraphBefore();
                oPara4.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                oPara4.BaseLineAlignment = WdBaselineAlignment.wdBaselineAlignAuto;
                oPara4.Range.Text = no.ToString() + ") " + error;
                oPara4.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                oPara4.Range.Font.Size = 11;
                oPara4.Range.Font.Color = WdColor.wdColorAutomatic;
                oPara4.Range.InsertParagraphAfter();
            }
            fname.Close();
            //********To export pdf****************

            string ver = APP_ID.Substring(APP_ID.Length - 2);
            try
            {
                string paramExportFilePath = @"D:\MCD\Report\" + ID + "_" + ver + "_ValidationReport.PDF";
                string paramExportFilePath2 = @"D:\MCD\Report\" + Dwgreader.GetValue(1).ToString() + "-" + ID + "_" + ver + "_ValidationReport.PDF";
                WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
                bool paramOpenAfterExport = false;
                WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                doc.ExportAsFixedFormat(paramExportFilePath,
                    paramExportFormat, paramOpenAfterExport,
                    paramExportOptimizeFor, paramExportRange, paramStartPage,
                    paramEndPage, paramExportItem, paramIncludeDocProps,
                    paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                    paramBitmapMissingFonts, paramUseISO19005_1,
                    ref missing);
                System.IO.File.Copy(paramExportFilePath, paramExportFilePath2, true);
                switch (fid)
                {
                    case 1:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport.PDF", true);
                        break;
                    case 2:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_CC.PDF", true);
                        break;
                    case 3:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_Revised.PDF", true);
                        break;
                    case 4:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_Regularized.PDF", true);
                        break;
                    case 5:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_AA.PDF", true);
                        break;
                    case 6:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_REVDN.PDF", true);
                        break;
                    case 7:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_SARAL_Revise.PDF", true);
                        break;
                    case 8:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_SANCTION_Up_To_500_Sqmt.PDF", true);
                        break;
                    case 9:
                        System.IO.File.Copy(paramExportFilePath, "D:\\To-Erp\\" + ID + "_" + ver + "_ValidationReport_Revised_SANCTION_Up_To_500_Sqmt.PDF", true);
                        break;
                }

            }
            catch
            {
                object DocFilename = @"D:\MCD\Report\" + Dwgreader.GetValue(1).ToString() + "-" + ID + "_" + ver + "_ValidationReport.DOC"; ;
                doc.SaveAs(ref DocFilename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            //doc.Close(ref savechanges, ref  missing, ref missing);
            WordApp.Quit(ref savechanges, ref  missing, ref missing);
        }
    }
}
