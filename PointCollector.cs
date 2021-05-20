using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using ACAD = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyAuto
{
    class PointCollector : IDisposable
    {

        public enum Shape
        {
            Window,
            Fence,
            Polygon,
            RegularPolygon,
            Circle,
        }

        private Shape mShape;
        private Autodesk.AutoCAD.DatabaseServices.Polyline mTempPline;
        private Point3d m1stPoint; //Center Point
        private double mDist;
        private int mSegmentCount = 36;

        public Point3dCollection CollectedPoints { get; set; }

        public PointCollector(Shape shape)
        {
            mShape = shape;

            CollectedPoints = new Point3dCollection();
        }

        private void Editor_PointMonitor(object sender, PointMonitorEventArgs e)
        {
            if (mTempPline != null)
                TransientManager.CurrentTransientManager.EraseTransient(mTempPline, new IntegerCollection());

            Point3d compPt = e.Context.ComputedPoint.TransformBy(ACAD.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem.Inverse());

            if (mShape == Shape.RegularPolygon)
            {
                BuildupRegularPolygonVertices(compPt);
            }
            else if (mShape == Shape.Circle)
            {
                BuildupRegularPolygonVertices(compPt);
            }

            TransientManager.CurrentTransientManager.AddTransient(mTempPline, TransientDrawingMode.Main, 0, new IntegerCollection());
        }

        public void Collect()
        {
            if (mShape == Shape.RegularPolygon)
            {
                PromptIntegerResult prPntRes;
                PromptIntegerOptions prPntOpt = new PromptIntegerOptions("");
                prPntOpt.AllowNone = true;
                prPntOpt.AllowArbitraryInput = true;
                prPntOpt.AllowNegative = false;
                prPntOpt.AllowZero = false;
                prPntOpt.DefaultValue = 5;
                prPntOpt.LowerLimit = 3;
                prPntOpt.Message = "\nRegular polygon side";
                prPntOpt.UpperLimit = 36;
                prPntOpt.UseDefaultValue = true;

                prPntRes = ACAD.DocumentManager.MdiActiveDocument.Editor.GetInteger(prPntOpt);
                if (prPntRes.Status == PromptStatus.OK)
                {
                    mSegmentCount = prPntRes.Value;
                }
                else
                {
                    throw new System.Exception("Regular polygon side input failed!");
                }

                CollectRegularPolygonPoints();
            }
            else if (mShape == Shape.Circle)
            {
                CollectRegularPolygonPoints();
            }
        }

        private void BuildupRegularPolygonVertices(Point3d tempPt)
        {
            if (mTempPline != null && !mTempPline.IsDisposed)
            {
                mTempPline.Dispose();
                mTempPline = null;
            }

            mTempPline = new Autodesk.AutoCAD.DatabaseServices.Polyline();
            mTempPline.SetDatabaseDefaults();
            mTempPline.Closed = true;
            mTempPline.ColorIndex = 7;

            mDist = m1stPoint.DistanceTo(tempPt);
            double angle = m1stPoint.GetVectorTo(tempPt).AngleOnPlane(new Plane(Point3d.Origin, Vector3d.ZAxis));
            CollectedPoints.Clear();
            for (int i = 0; i < mSegmentCount; i++)
            {
                Point3d pt = m1stPoint.Add(new Vector3d(mDist * (Math.Cos(angle + Math.PI * 2 * i / mSegmentCount)),
                                                        mDist * (Math.Sin(angle + Math.PI * 2 * i / mSegmentCount)),
                                                        m1stPoint.Z));
                CollectedPoints.Add(pt);
                mTempPline.AddVertexAt(mTempPline.NumberOfVertices, new Point2d(pt.X, pt.Y), 0, 1, 1);
            }

            mTempPline.TransformBy(ACAD.DocumentManager.MdiActiveDocument.Editor.CurrentUserCoordinateSystem);
        }

        private void CollectRegularPolygonPoints()
        {
            PromptPointResult prPntRes1;
            PromptPointOptions prPntOpt = new PromptPointOptions("\nCenter");
            prPntOpt.AllowNone = true;

            prPntRes1 = ACAD.DocumentManager.MdiActiveDocument.Editor.GetPoint(prPntOpt);
            if (prPntRes1.Status == PromptStatus.OK)
            {
                m1stPoint = prPntRes1.Value;
            }
            else
            {
                throw new System.Exception("Center picking failed!");
            }

            ACAD.DocumentManager.MdiActiveDocument.Editor.PointMonitor += Editor_PointMonitor;

            PromptDistanceOptions prPntOpt2 = new PromptDistanceOptions("");
            prPntOpt2.AllowArbitraryInput = true;
            prPntOpt2.AllowNegative = false;
            prPntOpt2.AllowNone = true;
            prPntOpt2.AllowZero = false;
            prPntOpt2.BasePoint = m1stPoint;
            prPntOpt2.DefaultValue = 10.0;
            prPntOpt2.Message = "\nRadius";
            prPntOpt2.Only2d = true;
            prPntOpt2.UseBasePoint = true;
            prPntOpt2.UseDashedLine = true;
            prPntOpt2.UseDefaultValue = true;

            PromptDoubleResult prPntRes2 = ACAD.DocumentManager.MdiActiveDocument.Editor.GetDistance(prPntOpt2);
            if (prPntRes2.Status != PromptStatus.OK)
                throw new System.Exception("Radius input failed!");

            mDist = prPntRes2.Value;

            ACAD.DocumentManager.MdiActiveDocument.Editor.PointMonitor -= Editor_PointMonitor;
        }

 
        public void Dispose()
        {
            ACAD.DocumentManager.MdiActiveDocument.Editor.PointMonitor -= Editor_PointMonitor;
            if (mTempPline != null)
                TransientManager.CurrentTransientManager.EraseTransient(mTempPline, new IntegerCollection());

            if (mTempPline != null && !mTempPline.IsDisposed)
                mTempPline.Dispose();

            CollectedPoints.Dispose();
        }
    }
}

  

