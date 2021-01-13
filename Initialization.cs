using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Microsoft.Office.Interop.Excel;

namespace MyAuto
{
    public class Initialization
    {
        [CommandMethod("Insert_Block_usingexcel")] public void Insert_Block_usingexcel()
        {
            string _path = "D:\\Videos & docs\\CAD\\AutoCAD\\Lynda AutoCAD Add-ins Using VB.NET_git.ir\\Exercise Files\\MyAuto\\Output.xlsx";
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            myExcel.Visible = false;
            var myWB = myExcel.Workbooks.Open(_path);
            var myWS = myExcel.ActiveSheet;
            int lastUsedRow = Convert.ToInt32(myWS.cells(1, "M").value);

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction myTrans = db.TransactionManager.StartTransaction())
            {
                BlockTable MyBT = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);

                for (int RowN = 3; RowN <= lastUsedRow; RowN++)
                {
                    //Get the block definition "Check"
                    string blockName = myWS.cells(RowN, "B").value;
                    if (blockName == null) { continue; }

                    var Xpo = myWS.cells(RowN, "C").value;
                    var Ypo = myWS.cells(RowN, "D").value;
                    var Zpo = myWS.cells(RowN, "E").value;
                    Point3d point = new Point3d(Xpo, Ypo, Zpo);

                    var X_Scale = myWS.cells(RowN, "F").value;
                    var Y_Scale = myWS.cells(RowN, "G").value;
                    var Z_Scale = myWS.cells(RowN, "H").value;
                    Scale3d scale = new Scale3d(X_Scale, Y_Scale, Z_Scale);
                    Double Rot = myWS.cells(RowN, "I").value;

                    //Create new BlockReference, and link it to our block definition
                    BlockTableRecord blockDef = (BlockTableRecord)MyBT[blockName].GetObject(OpenMode.ForRead);
                    //Also open modelspace - we'll be adding our BlockReference to it
                    BlockTableRecord curSpace = (BlockTableRecord)myTrans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);

                    if (myWS.cells(RowN + 1, "B").value == null)
                    {
                        Dictionary<string, string> attValues = new Dictionary<string, string>();
                        for (int i = RowN + 1; i <= RowN + 4; i++) //last cells come here and have error
                        {
                            var Att1 = myWS.cells(i, "J").value.ToString();
                            var Att2 = myWS.cells(i, "K").value.ToString();
                            attValues.Add(Att1, Att2);
                        }

                        BlockReference br = curSpace.InsertBlockReference(blockName, point, scale, Rot, attValues);
                    }
                    else
                    {
                        using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                        {
                            //Add Scale factor and rotation of block
                            blockRef.ScaleFactors = scale;
                            blockRef.Rotation = Rot;

                            //Add the block reference to modelspace
                            curSpace.AppendEntity(blockRef);
                            myTrans.AddNewlyCreatedDBObject(blockRef, true);
                        }
                    }

                }
                myTrans.Commit();
            }
            ed.WriteMessage("Blocks has been inserted");
            myWB.Close();
            myWS = null;
            myWB = null;
            myExcel = null;
        }

        [CommandMethod("InsertDWGblock")] public void InsertDWGblock()
        {
            const string filename = @"F:\gile\TestDrawing.dwg";
            Scale3d Scale = new Scale3d(1, 1, 1);

            using (Database db = new Database(false, true))
            {
                db.ReadDwgFile(filename, System.IO.FileShare.ReadWrite, false, null);
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    ObjectId btrId = bt.GetBlock(@"F:\gile\Gile_blocs\bloc-att.dwg");
                    if (btrId == ObjectId.Null)
                        return;
                    Dictionary<string, string> attValues = new Dictionary<string, string>();
                    attValues.Add("ATT1", "foo");
                    attValues.Add("ATT2", "bar");
                    BlockTableRecord curSpace = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    BlockReference br = curSpace.InsertBlockReference("bloc-att", Point3d.Origin, Scale, 0, attValues);
                    tr.Commit();
                }
                db.SaveAs(filename, DwgVersion.Current);
            }
        }

        [CommandMethod("AddBlockTest")] static public void AddBlockTest()
        {

            Database db = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            using (Transaction myTrans = db.TransactionManager.StartTransaction())
            {
                //Get the block definition "Check".
                string blockName = "BOMLINE";

                BlockTable MyBT = (BlockTable) db.BlockTableId.GetObject(OpenMode.ForRead);
                BlockTableRecord blockDef = (BlockTableRecord) MyBT[blockName].GetObject(OpenMode.ForRead);

                //Also open modelspace - we'll be adding our BlockReference to it
                BlockTableRecord ms = (BlockTableRecord) MyBT[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite);

                //Create new BlockReference, and link it to our block definition
                Point3d point = new Point3d(2.0, 4.0, 6.0);

                using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                {
                    //Add the block reference to modelspace
                    ms.AppendEntity(blockRef);
                    myTrans.AddNewlyCreatedDBObject(blockRef, true);

                    //Iterate block definition to find all non-constant // AttributeDefinitions
                    foreach (ObjectId id in blockDef)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition attDef = obj as AttributeDefinition;

                        if ((attDef != null) && (!attDef.Constant))
                        {
                            //This is a non-constant AttributeDefinition //Create a new AttributeReference
                            using (AttributeReference myAttRef = new AttributeReference())
                            {
                                myAttRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                myAttRef.TextString = "Hello World";
                                //Add the AttributeReference to the BlockReference
                                blockRef.AttributeCollection.AppendAttribute(myAttRef);
                                myTrans.AddNewlyCreatedDBObject(myAttRef, true);
                            }
                        }
                    }
                }
                myTrans.Commit(); //Our work here is done
            } //Trans Finish
        }

        [CommandMethod("ExportBlockDetails_Excel")] public void ExportBlockDetails_Excel()  
        { //Delete all blocks(like that line....)
            Microsoft.Office.Interop.Excel.Application myExcel = new Microsoft.Office.Interop.Excel.Application();
            myExcel.Visible = false;
            var myWB = myExcel.Workbooks.Add();
            var myWS = myExcel.ActiveSheet;
            Database db = HostApplicationServices.WorkingDatabase;
            int curRow = 1;
            myWS.cells(curRow, "A").value = HostApplicationServices.WorkingDatabase.OriginalFileName;
            curRow += 1;
            foreach (string myName in BlockUtils.GetBlockNames(db))
            {
                foreach (ObjectId myBrefID in BlockUtils.GetBlockIDs(db, myName))
                {
                    myWS.cells(curRow, "A").value = myBrefID.ToString();
                    myWS.cells(curRow, "B").value = myName;

                    using (Transaction myTrans = myBrefID.Database.TransactionManager.StartTransaction())
                    {
                        BlockReference myBref = (BlockReference)myBrefID.GetObject(OpenMode.ForRead);
                        myWS.cells(curRow, "C").value = myBref.Position.X;
                        myWS.cells(curRow, "D").value = myBref.Position.Y;
                        myWS.cells(curRow, "E").value = myBref.Position.Y;
                        myWS.cells(curRow, "F").value = myBref.ScaleFactors.X;
                        myWS.cells(curRow, "G").value = myBref.ScaleFactors.Y;
                        myWS.cells(curRow, "H").value = myBref.ScaleFactors.Y;
                        myWS.cells(curRow, "I").value = myBref.Rotation;
                    }
                    curRow += 1;

                    foreach (KeyValuePair<string, string> myKVP in BlockUtils.GetAttributes(myBrefID))
                    {
                        myWS.cells(curRow, "J").value = myKVP.Key;
                        myWS.cells(curRow, "K").value = myKVP.Value;
                        curRow += 1;
                    }
                }
            }

            string filename = "C:\\temp\\Output.xlsx";

            myWB.SaveAs(filename, XlFileFormat.xlWorkbookDefault, Type.Missing, 
                Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, 
                Type.Missing, Type.Missing);

            myWS = null;
            myWB = null;
            myExcel = null;
        }

        [CommandMethod("ExportBlockDetails_Txt")] public void ExportBlockDetails_Txt()
        {
            System.IO.FileInfo myFIO = new System.IO.FileInfo("C:\\temp\\blocks.txt");
            if (myFIO.Directory.Exists == false)
            {
                myFIO.Directory.Create();
            }
            Database dbToUse = HostApplicationServices.WorkingDatabase;
            System.IO.StreamWriter mySW = new System.IO.StreamWriter(myFIO.FullName);
            mySW.WriteLine(HostApplicationServices.WorkingDatabase.Filename);
            foreach (string myName in BlockUtils.GetBlockNames(dbToUse))
            {
                foreach (ObjectId myBrefID in BlockUtils.GetBlockIDs(dbToUse, myName))
                {
                    mySW.WriteLine(" " + myName);
                    foreach (KeyValuePair<string, string> myKVP in BlockUtils.GetAttributes(myBrefID))
                    {
                        mySW.WriteLine(" " + myKVP.Key + " " + myKVP.Value);
                    }
                }
            }
            mySW.Close();
            mySW.Dispose();
        }

        [CommandMethod("MergeBlocks")] public static void MergeBlocks()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Get the name of the first block to merge
            PromptResult pr = ed.GetString("\nEnter name of first block");
            if (pr.Status != PromptStatus.OK) return;
            string first = pr.StringResult.ToUpper();

            using (Transaction myTrans = doc.TransactionManager.StartTransaction())
            {
                BlockTable MyBT = (BlockTable)db.BlockTableId.GetObject(OpenMode.ForRead);
                //var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                // Check whether the first block exists
                if (MyBT.Has(first))
                {
                    // Get the name of the second block to merge
                    pr = ed.GetString("\nEnter name of second block");
                    if (pr.Status != PromptStatus.OK) return;
                    string second = pr.StringResult.ToUpper();

                    // Check whether the second block exists
                    if (MyBT.Has(second))
                    {
                        // Get the name of the new block
                        pr = ed.GetString("\nEnter name for new block");
                        if (pr.Status != PromptStatus.OK) return;
                        string merged = pr.StringResult.ToUpper();

                        // Make sure the new block doesn't already exist
                        if (!MyBT.Has(merged))
                        {
                            // We need to collect the contents of the two blocks
                            var ids = new ObjectIdCollection();

                            // Open the two blocks to be merged
                            BlockTableRecord btr1 = myTrans.GetObject(MyBT[first], OpenMode.ForRead) as BlockTableRecord;
                            BlockTableRecord btr2 = myTrans.GetObject(MyBT[second], OpenMode.ForRead) as BlockTableRecord;

                            // Use LINQ to get IEnumerable<ObjectId> for the blocks
                            var en1 = btr1.Cast<ObjectId>();
                            var en2 = btr2.Cast<ObjectId>();

                            // Add the complete contents to our collection
                            // (we could also apply some filtering, here, such as making 
                            // sure we only include attributes with the same name once)
                            var tt1 = en1.ToArray<ObjectId>();
                            var tt2 = en2.ToArray<ObjectId>();
                            for(int i =0; i < tt1.Length; i++) { ids.Add(tt1[i]); }
                            for(int i =0; i < tt2.Length; i++) { ids.Add(tt2[i]); }

                            // Create a new block table record for our merged block
                            var btr = new BlockTableRecord();
                            btr.Name = merged;
                            // Add it to the block table and the transaction
                            MyBT.UpgradeOpen();
                            var btrId = MyBT.Add(btr);
                            myTrans.AddNewlyCreatedDBObject(btr, true);

                            // Deep clone the contents of our two blocks into the new one
                            var idMap = new IdMapping();
                            db.DeepCloneObjects(ids, btrId, idMap, false);

                            ed.WriteMessage("\nBlock \"{0}\" created.", merged);
                        }
                        else { ed.WriteMessage("\nDrawing already contains a block named \"{0}\".", merged); }
                    }
                    else { ed.WriteMessage("\nBlock \"{0}\" not found.", second); }
                }
                else { ed.WriteMessage("\nBlock \"{0}\" not found.", first); }
                // Always commit the transaction
                myTrans.Commit();
            }
        }

    }
}
