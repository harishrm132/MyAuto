using Autodesk.AutoCAD.DatabaseServices;
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
    }
}