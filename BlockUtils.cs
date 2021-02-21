using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using AcRx = Autodesk.AutoCAD.Runtime;

namespace MyAuto
{
    static class BlockUtils
    {
        public static ObjectId GetBlock(this BlockTable blockTable, string blockName)
        {
            if (blockTable == null)
                throw new ArgumentNullException("blockTable");

            Database db = blockTable.Database;
            if (blockTable.Has(blockName))
                return blockTable[blockName];

            try
            {
                string ext = Path.GetExtension(blockName);
                if (ext == "")
                    blockName += ".dwg";
                string blockPath;
                if (File.Exists(blockName))
                    blockPath = blockName;
                else
                    blockPath = HostApplicationServices.Current.FindFile(blockName, db, FindFileHint.Default);

                blockTable.UpgradeOpen();
                using (Database tmpDb = new Database(false, true))
                {
                    tmpDb.ReadDwgFile(blockPath, FileShare.Read, true, null);
                    return blockTable.Database.Insert(Path.GetFileNameWithoutExtension(blockName), tmpDb, true);
                }
            }
            catch
            {
                return ObjectId.Null;
            }
        }

        public static BlockReference InsertBlockReference(this BlockTableRecord target, string blkName, Point3d insertPoint, Scale3d scale, Double Rot,
            Dictionary<string, string> attValues = null)
        {
            if (target == null)
                throw new ArgumentNullException("target");

            Database db = target.Database;
            Transaction tr = db.TransactionManager.TopTransaction;
            if (tr == null)
                throw new AcRx.Exception(ErrorStatus.NoActiveTransactions);

            BlockReference br = null;
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            //ObjectId btrId = bt.GetBlock(blkName);
            BlockTableRecord blockDef = (BlockTableRecord)bt[blkName].GetObject(OpenMode.ForRead);
            ObjectId btrId = blockDef.ObjectId;

            if (btrId != ObjectId.Null)
            {
                //Create Block reference and Add Scale factor and rotation of block
                br = new BlockReference(insertPoint, btrId);
                br.ScaleFactors = scale;
                br.Rotation = Rot;

                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);
                target.AppendEntity(br);
                tr.AddNewlyCreatedDBObject(br, true);

                br.AddAttributeReferences(attValues);
            }
            return br;
        }

        public static void AddAttributeReferences(this BlockReference target, Dictionary<string, string> attValues)
        {
            if (target == null)
                throw new ArgumentNullException("target");

            Transaction tr = target.Database.TransactionManager.TopTransaction;
            if (tr == null)
                throw new AcRx.Exception(ErrorStatus.NoActiveTransactions);

            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(target.BlockTableRecord, OpenMode.ForRead);
            RXClass attDefClass = RXClass.GetClass(typeof(AttributeDefinition));

            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass != attDefClass)
                    continue;
                AttributeDefinition attDef = (AttributeDefinition)tr.GetObject(id, OpenMode.ForRead);
                AttributeReference attRef = new AttributeReference();
                attRef.SetAttributeFromBlock(attDef, target.BlockTransform);
                if (attValues != null && attValues.ContainsKey(attDef.Tag.ToUpper()))
                {
                    attRef.TextString = attValues[attDef.Tag.ToUpper()];
                }
                target.AttributeCollection.AppendAttribute(attRef);
                tr.AddNewlyCreatedDBObject(attRef, true);
            }
        }

        public static void DisplayDynBlockProperties(Editor ed, BlockReference br, string name)
        {
            if (br != null && br.IsDynamicBlock)
            {
                ed.WriteMessage("\nDynamic properties for \"{0}\"\n", name);
                // Get the dynamic block's property collection
                DynamicBlockReferencePropertyCollection pc = br.DynamicBlockReferencePropertyCollection;
                foreach (DynamicBlockReferenceProperty prop in pc)
                {
                    //Start with property name, type and description
                    ed.WriteMessage("\n Property: \"{0}\" : {1}", prop.PropertyName, prop.UnitsType);
                    if (prop.Description != "") { ed.WriteMessage("\n Description: {0}", prop.Description); }
                    if (prop.ReadOnly) { ed.WriteMessage("\n (Read Only) "); }
                    // Get the allowed values, if it's constrained
                    bool IsFirst = true;
                    foreach (object value in prop.GetAllowedValues())
                    {
                        ed.WriteMessage((IsFirst ? "\n  Allowed values: [" : ", "));
                        ed.WriteMessage("\"{0}\"", value);
                        IsFirst = false;
                    }
                    if (!IsFirst) { ed.WriteMessage("]"); }
                    ed.WriteMessage("\n current value: \"{0}\"\n");
                }
            }
        }

        public static List<string> GetBlockNames(Database DBIn)
        {
            List<string> retList = new List<string>();
            using (Transaction myTrans = DBIn.TransactionManager.StartTransaction())
            {
                BlockTable myBT = (BlockTable)DBIn.BlockTableId.GetObject(OpenMode.ForRead);
                foreach (ObjectId myOID in myBT)
                {
                    BlockTableRecord myBTR = (BlockTableRecord)myOID.GetObject(OpenMode.ForRead);
                    if (myBTR.IsLayout == false | myBTR.IsAnonymous == false)
                    {
                        retList.Add(myBTR.Name);
                    }
                }
            }
            return (retList);
        }

        public static ObjectIdCollection GetBlockIDs(Database DBIn, string BlockName)
        {
            ObjectIdCollection retCollection = new ObjectIdCollection();
            using (Transaction myTrans = DBIn.TransactionManager.StartTransaction())
            {
                BlockTable myBT = (BlockTable)DBIn.BlockTableId.GetObject(OpenMode.ForRead);
                if (myBT.Has(BlockName))
                {
                    BlockTableRecord myBTR = (BlockTableRecord)myBT[BlockName].GetObject(OpenMode.ForRead);
                    retCollection = (ObjectIdCollection)myBTR.GetBlockReferenceIds(true, true);
                    myTrans.Commit();
                    return (retCollection);
                }
                else
                {
                    myTrans.Commit();
                    return (retCollection);
                }
            }
        }

        public static Dictionary<string, string> GetAttributes(ObjectId BlockRefID)
        {
            Dictionary<string, string> retDictionary = new Dictionary<string, string>();
            using (Transaction myTrans = BlockRefID.Database.TransactionManager.StartTransaction())
            {
                BlockReference myBref = (BlockReference)BlockRefID.GetObject(OpenMode.ForRead);
                if (myBref.AttributeCollection.Count == 0)
                {
                    return (retDictionary);
                }
                else
                {
                    foreach (ObjectId myBRefID in myBref.AttributeCollection)
                    {
                        AttributeReference myAttRef = (AttributeReference)myBRefID.GetObject(OpenMode.ForRead);
                        if (retDictionary.ContainsKey(myAttRef.Tag) == false)
                        {
                            retDictionary.Add(myAttRef.Tag, myAttRef.TextString); //prompt
                        }
                    }
                    return (retDictionary);
                }
            }
        }

        public static Dictionary<double, double> GetCoordinates(ObjectId BlockRefID)
        {
            Dictionary<double, double> retDictionary = new Dictionary<double, double>();
            using (Transaction myTrans = BlockRefID.Database.TransactionManager.StartTransaction())
            {
                BlockReference myBref = (BlockReference)BlockRefID.GetObject(OpenMode.ForRead);
                retDictionary.Add(myBref.Position.X, myBref.Position.Y);
                return retDictionary;
            }
        }

    }
}