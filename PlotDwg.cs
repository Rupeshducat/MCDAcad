using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD .PlottingServices;

namespace MCD
{
    class PlotDwg
    {

        public static void PlotCurrentLayout(string FileName, string MediaName)
        {
            MediaName = "ISO_A0_(841.00_x_1189.00_MM)";
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr;
                acLayoutMgr = LayoutManager.Current;

                // Get the current layout and output its name in the Command Line window
                Layout acLayout;
                acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout),
                                                   OpenMode.ForRead) as Layout;

                // Get the PlotInfo from the layout
                PlotInfo acPlInfo = new PlotInfo();
                acPlInfo.Layout = acLayout.ObjectId;

                // Get a copy of the PlotSettings from the layout
                PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
                acPlSet.CopyFrom(acLayout);

                // Update the PlotSettings object
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;

                // Set the plot type
                acPlSetVdr.SetPlotType(acPlSet,
                                             Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);

                // Set the plot scale
                acPlSetVdr.SetUseStandardScale(acPlSet, true);
                acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);

                // Center the plot
                acPlSetVdr.SetPlotCentered(acPlSet, true);

                // Set the plot device to use
                acPlSetVdr.SetPlotConfigurationName(acPlSet, "DWG To PDF.pc3",
                                                          MediaName);//"ISO_A0_(841.00_x_1189.00_MM)"

                // Set the plot info as an override since it will
                // not be saved back to the layout
                acPlInfo.OverrideSettings = acPlSet;

                // Validate the plot info
                PlotInfoValidator acPlInfoVdr = new PlotInfoValidator();
                acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                acPlInfoVdr.Validate(acPlInfo);

                // Check to see if a plot is already in progress
                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                {
                    using (PlotEngine acPlEng = PlotFactory.CreatePublishEngine())
                    {
                        // Track the plot progress with a Progress dialog
                        PlotProgressDialog acPlProgDlg = new PlotProgressDialog(false,
                                                                                1,
                                                                                true);

                        using (acPlProgDlg)
                        {
                            // Define the status messages to display when plotting starts
                            //acPlProgDlg.set_PlotMsgString(PlotMessageIndex.DialogTitle,
                            //                              "Plot Progress");

                            //acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage,
                            //                              "Cancel Job");

                            //acPlProgDlg.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage,
                            //                              "Cancel Sheet");

                            //acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption,
                            //                              "Sheet Set Progress");

                            //acPlProgDlg.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption,
                            //                              "Sheet Progress");

                            // Set the plot progress range
                            acPlProgDlg.LowerPlotProgressRange = 0;
                            acPlProgDlg.UpperPlotProgressRange = 100;
                            acPlProgDlg.PlotProgressPos = 0;

                            // Display the Progress dialog
                            acPlProgDlg.OnBeginPlot();
                            acPlProgDlg.IsVisible = false;

                            // Start to plot the layout
                            acPlEng.BeginPlot(acPlProgDlg, null);

                            // Define the plot output
                            acPlEng.BeginDocument(acPlInfo,
                                                  acDoc.Name,
                                                  null,
                                                  1,
                                                  true,
                                                  FileName);

                            // Display information about the current plot
                            acPlProgDlg.set_PlotMsgString(PlotMessageIndex.Status,
                                                          "Plotting: " + acDoc.Name + " - " +
                                                          acLayout.LayoutName);

                            // Set the sheet progress range
                            acPlProgDlg.OnBeginSheet();
                            acPlProgDlg.LowerSheetProgressRange = 0;
                            acPlProgDlg.UpperSheetProgressRange = 100;
                            acPlProgDlg.SheetProgressPos = 0;

                            // Plot the first sheet/layout
                            PlotPageInfo acPlPageInfo = new PlotPageInfo();
                            acPlEng.BeginPage(acPlPageInfo,
                                              acPlInfo,
                                              true,
                                              null);

                            acPlEng.BeginGenerateGraphics(null);
                            acPlEng.EndGenerateGraphics(null);

                            // Finish plotting the sheet/layout
                            acPlEng.EndPage(null);
                            acPlProgDlg.SheetProgressPos = 100;
                            acPlProgDlg.OnEndSheet();

                            // Finish plotting the document
                            acPlEng.EndDocument(null);

                            // Finish the plot
                            acPlProgDlg.PlotProgressPos = 100;
                            acPlProgDlg.OnEndPlot();
                            acPlEng.EndPlot(null);
                        }
                    }
                }
            }
        }
    }
}
