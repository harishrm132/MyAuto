using ACAD = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyAuto
{
    class Selectionsets
    {
        //• PICKFIRST system variable must be set to 1
        //• UsePickSet command flag must be defined with the command that should use the Pickfirst selection set
        //• Call the SelectImplied method to obtain the PickFirst selection set
        //Call GetSelection method - Prompts the user to pick objects from the screen.
        [CommandMethod("CheckSelectionSets", CommandFlags.UsePickSet)] public static void SelectSet()
        {
            Editor ed = ACAD.DocumentManager.MdiActiveDocument.Editor;
            SelectionSet acSSet1;
            PromptSelectionResult acSSPrompt = ed.SelectImplied();
            // Clear the PickFirst selection set
            ObjectId[] idarrayEmpty = new ObjectId[0];
            ed.SetImpliedSelection(idarrayEmpty);
            // Request for objects to be selected in the drawing area
            acSSPrompt = ed.GetSelection(); 

            //selects the objects within and that intersect a crossing window.
            acSSPrompt = ed.SelectCrossingWindow(new Point3d(2, 2, 0), new Point3d(10, 8, 0));

            //Add To or Merge Multiple Selection Sets
            acSSet1 = acSSPrompt.Value;
            ObjectIdCollection coll = new ObjectIdCollection(acSSet1.GetObjectIds());
            SelectionSet acSSet2 = acSSPrompt.Value;
            if (coll.Count == 0)
            {
                coll = new ObjectIdCollection(acSSet2.GetObjectIds());  
            }
            else
            {
                foreach (ObjectId id in acSSet2.GetObjectIds())
                {
                    coll.Add(id);
                }
            }

            //Define Rules for Selection Filters
                // A selection filter list can be used to filter selected objects by properties or type.
        }
    }
}
