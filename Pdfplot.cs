using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Internal.Reactors;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using ACAD = Autodesk.AutoCAD.ApplicationServices.Application;
using ED = Autodesk.AutoCAD.EditorInput.Editor;
namespace MyAuto
{
    public class Pdfplot : IExtensionApplication
    {
        [DllImport("accore.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "acedTrans")]
        static extern int acedTrans(double[] point, IntPtr fromRb, IntPtr toRb, int disp, double[] result);
        List<string> msgs = new List<string>();

        public class Blk2Plt
        {
            public BlockReference BlockRef;
            public Layout LayoutObj;
        }

        public void Terminate()
        {
            ApplicationEventManager cadWinEvnts = Autodesk.AutoCAD.Internal.Reactors.ApplicationEventManager.Instance();
            msgs.Sort();
            ED ed = ACAD.DocumentManager.MdiActiveDocument.Editor;
            foreach (string msg in msgs)
            {
                ed.WriteMessage(msg);
            }
        }

        public void Initialize()
        {

        }

        [CommandMethod("PlotToPdf")] static public void PlotToPdf()
        {
            String BlockName = "SHEETGOST";
            String PrinterName = "DWG To PDF.pc3";
            String PaperSize = "ISO_A4_(210.00_x_297.00_MM)";
            //String PaperSize = "ISO_A4_(297.00_x_210.00_MM)";
            //String PaperSize = "ISO_A3_(297.00_x_420.00_MM)";
            String OutPath = "c:\\temp\\plot2pdf";

            PlotBlockColToPDF(BlockName, PrinterName, OutPath, PaperSize);
        }

        static void PlotBlockColToPDF(String BlockName, String PrinterName, String OutPath, String PaperSize = "ISO_A4_(297.00_x_210.00_MM)")
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            Transaction tr = db.TransactionManager.StartTransaction();

            Object SysVarBackPlot = Application.GetSystemVariable("BACKGROUNDPLOT");
            Application.SetSystemVariable("BACKGROUNDPLOT", 0);

            using (tr)
            {
                List<Blk2Plt> BlocksToPlot = new List<Blk2Plt>();
                GetBlocksToPlotCollection(db, ed, tr, BlockName, ref BlocksToPlot);//Getting collection of blocks
                ed.WriteMessage("\nThe number of blocks found: " + BlocksToPlot.Count + "\n\n");
                if (BlocksToPlot.Count < 1) return;

                if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                {
                    PlotEngine pe = PlotFactory.CreatePublishEngine();
                    using (pe)
                    {
                        PlotProgressDialog ppd = new PlotProgressDialog(false, BlocksToPlot.Count, true);
                        using (ppd)
                        {
                            int numSheet = 1;
                            // Setting up the PlotProgress dialog
                            ppd.set_PlotMsgString(PlotMessageIndex.DialogTitle, "Custom Plot Progress");
                            ppd.set_PlotMsgString(PlotMessageIndex.CancelJobButtonMessage, "Cancel Job");
                            ppd.set_PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage, "Cancel Sheet");
                            ppd.set_PlotMsgString(PlotMessageIndex.SheetSetProgressCaption, "Sheet Set Progress");
                            ppd.set_PlotMsgString(PlotMessageIndex.SheetProgressCaption, "Sheet Progress");
                            ppd.LowerPlotProgressRange = 0;
                            ppd.UpperPlotProgressRange = 100;
                            ppd.PlotProgressPos = 0;
                            ppd.OnBeginPlot();
                            ppd.IsVisible = true;

                            pe.BeginPlot(ppd, null);

                            foreach (Blk2Plt gblk in BlocksToPlot)
                            {
                                // Starting new page
                                ppd.StatusMsgString = "Plotting block " + numSheet.ToString() + " of " + BlocksToPlot.Count.ToString();
                                ppd.OnBeginSheet();
                                ppd.LowerSheetProgressRange = 0;
                                ppd.UpperSheetProgressRange = 100;
                                ppd.SheetProgressPos = 0;

                                PlotInfoValidator piv = new PlotInfoValidator();
                                piv.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                PlotPageInfo ppi = new PlotPageInfo();
                                PlotInfo pi = new PlotInfo();
                                BlockReference blk = gblk.BlockRef;
                                Layout lo = gblk.LayoutObj;

                                // Getting coodinates of window to plot
                                Extents3d ext = (Extents3d)blk.Bounds;
                                Point3d first = ext.MaxPoint;
                                Point3d second = ext.MinPoint;
                                ResultBuffer rbFrom = new ResultBuffer(new TypedValue(5003, 1)), rbTo = new ResultBuffer(new TypedValue(5003, 2));
                                double[] firres = new double[] { 0, 0, 0 };
                                double[] secres = new double[] { 0, 0, 0 };
                                acedTrans(first.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, firres);
                                acedTrans(second.ToArray(), rbFrom.UnmanagedObject, rbTo.UnmanagedObject, 0, secres);
                                Extents2d window = new Extents2d(firres[0], firres[1], secres[0], secres[1]);

                                // We need a PlotSettings object based on the layout settings which we then customize
                                PlotSettings ps = new PlotSettings(lo.ModelType);
                                LayoutManager.Current.CurrentLayout = lo.LayoutName;
                                pi.Layout = lo.Id;
                                ps.CopyFrom(lo);

                                // The PlotSettingsValidator helps create a valid PlotSettings object
                                PlotSettingsValidator psv = PlotSettingsValidator.Current;
                                psv.SetPlotWindowArea(ps, window);
                                psv.SetPlotType(ps, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                                psv.SetUseStandardScale(ps, true);
                                psv.SetStdScaleType(ps, StdScaleType.ScaleToFit);
                                psv.SetPlotCentered(ps, true);
                                psv.SetPlotConfigurationName(ps, PrinterName, PaperSize);
                                psv.SetZoomToPaperOnUpdate(ps, true);

                                pi.OverrideSettings = ps;
                                piv.Validate(pi);

                                if (numSheet == 1) pe.BeginDocument(pi, doc.Name, null, 1, true, OutPath); // Create document for the first page
                                    
                                // Plot the window
                                pe.BeginPage(ppi, pi, (numSheet == BlocksToPlot.Count), null);
                                pe.BeginGenerateGraphics(null);
                                ppd.SheetProgressPos = 50;
                                pe.EndGenerateGraphics(null);

                                // Finish the sheet
                                pe.EndPage(null);
                                ppd.SheetProgressPos = 100;
                                ppd.PlotProgressPos += Convert.ToInt32(100 / BlocksToPlot.Count);
                                ppd.OnEndSheet();
                                numSheet++;
                            }
                            // Finish the document and finish the plot
                            pe.EndDocument(null);
                            ppd.PlotProgressPos = 100;
                            ppd.OnEndPlot();
                            pe.EndPlot(null);
                            ed.WriteMessage("\nPlot completed successfully!\n\n");
                        }
                    }
                }
                else
                {
                    ed.WriteMessage("\nAnother plot is in progress.\n\n");
                }
                tr.Commit();
            }
            Application.SetSystemVariable("BACKGROUNDPLOT", SysVarBackPlot);
        }

        static void GetBlocksToPlotCollection(Database db, Editor ed, Transaction tr, String BlockName, ref List<Blk2Plt> BlockCol)
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId btrId in bt)
            {
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                if (btr.IsLayout)
                {
                    Layout lo = (Layout)tr.GetObject(btr.LayoutId, OpenMode.ForRead); 
                    BlockTableRecord ms = (BlockTableRecord)tr.GetObject(lo.BlockTableRecordId, OpenMode.ForRead);
                    foreach (ObjectId objId in ms)
                    {
                        Entity ent = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                        if (!ent.GetType().ToString().Contains("BlockReference")) continue;
                        BlockReference blk = (BlockReference)ent;
                        string Effn = EffectiveName(blk);
                        if (!Effn.ToUpper().Contains(BlockName.ToUpper())) continue;
                        Blk2Plt theBlk = new Blk2Plt();
                        theBlk.BlockRef = blk;
                        theBlk.LayoutObj = lo;
                        BlockCol.Add(theBlk);
                    }
                }
            }
        }

        static string EffectiveName(BlockReference blkref)
        {
            if (blkref.IsDynamicBlock)
            {
                using (BlockTableRecord obj = (BlockTableRecord)blkref.DynamicBlockTableRecord.GetObject(OpenMode.ForRead))
                    return obj.Name;
            }
            return blkref.Name;
        }


    }
}
