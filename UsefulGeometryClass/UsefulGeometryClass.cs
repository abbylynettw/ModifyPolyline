using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using TransactionManager = Autodesk.AutoCAD.ApplicationServices.TransactionManager;

namespace UsefulGeometryClass
{
    public class UsefulGeometryClass
    {
       
     

        // 计算两条曲线相交之后形成的边界线
        [CommandMethod("ABC")]
        public void CurveBoolean()
        {
            // 提示用户选择所要计算距离的两条直线
            ObjectIdCollection polyIds = new ObjectIdCollection();
            if (PromptSelectEnts("\n请选择要处理的复合线:", "LWPOLYLINE", ref polyIds))
            {
                Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
                if (polyIds.Count != 1)
                {
                    ed.WriteMessage("\n必须选择一条多段线进行操作.");
                    return;
                }

                var plineId = DrawWallMethod();
                Database db = HostApplicationServices.WorkingDatabase;
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    if (plineId.Equals(ObjectId.Null))
                    {
                        return;
                    }
                    polyIds.Add(plineId);
                    // 获得两条曲线的交点
                    Polyline poly1 = (Polyline)trans.GetObject(polyIds[0], OpenMode.ForRead);
                    Polyline poly2 = (Polyline)trans.GetObject(polyIds[1], OpenMode.ForRead);
                    Point3dCollection intPoints = new Point3dCollection();
                    poly1.IntersectWith(poly2, Intersect.OnBothOperands, intPoints, 0, 0);
                    if (intPoints.Count < 2)
                    {
                        ed.WriteMessage("\n曲线交点少于2个, 无法进行计算.");
                    }

                    // 根据交点和参数值获得交点之间的曲线
                    BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                    GetCurveBetweenIntPoints(trans, btr, poly1,poly2,intPoints);                   
                    trans.Commit();
                }
            }
        }

      

        // 提示用户选择一组实体
        public static bool PromptSelectEnts(string prompt, string entTypeFilter, ref ObjectIdCollection entIds)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptSelectionOptions pso = new PromptSelectionOptions {MessageForAdding = prompt};
            SelectionFilter sf = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, entTypeFilter) });
            PromptSelectionResult psr = ed.GetSelection(pso, sf);
            SelectionSet ss = psr.Value;
            if (ss != null)
            {
                entIds = new ObjectIdCollection(ss.GetObjectIds());
                return entIds.Count > 0;
            }
            else
            {
                return false;
            }
        }

        public ObjectId DrawWallMethod()
        {
          
          
            ObjectId plineid=ObjectId.Null;
            PromptPointResult start = Application.DocumentManager.MdiActiveDocument.Editor.GetPoint("请选择多段线的起点");
            if (start.Status != PromptStatus.OK) return ObjectId.Null;
            Polyline pline = new Polyline(); //第一条pl线
            using (var transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)transaction.GetObject(HostApplicationServices.WorkingDatabase.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                plineid = btr.AppendEntity(pline);
                transaction.AddNewlyCreatedDBObject(pline, true);
                pline.AddVertexAt(0, ToPoint2d(start.Value), 0, 0, 0);
                transaction.Commit();
            }
          
            Point3d tempCurrent = new Point3d(); //最后一次拖拽的点
            while (true)
            {
               var startPoint = pline.GetPoint3dAt(pline.NumberOfVertices - 1);
                EntityDrawing drawingJig = new EntityDrawing(startPoint, (currentPoint) =>
                {
                    tempCurrent = currentPoint; //每次拖拽为当前点的变量赋值
                    List<Entity> list = new List<Entity>();
                    list.Add(new Line(startPoint,currentPoint));
                    return list;
                }, "请选择下一点");
                PromptResult result1 = Application.DocumentManager.MdiActiveDocument.Editor.Drag(drawingJig);
               
                if (result1.Status == PromptStatus.OK)
                {
                    using (var trans = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction())
                    {
                        Polyline poly1 = (Polyline)trans.GetObject(plineid, OpenMode.ForWrite);
                        poly1.AddVertexAt(pline.NumberOfVertices, ToPoint2d(tempCurrent), 0, 0, 0);
                        poly1.DowngradeOpen();
                        trans.Commit();
                    }
                }
                else break;
            }
            return plineid;
        }

        // 获得多段线多个交点之间的子曲线，首尾两端删除
        private void GetCurveBetweenIntPoints(Transaction trans, BlockTableRecord btr, Polyline poly1,Polyline poly2,Point3dCollection points)
        {
            DBObjectCollection curves2 = poly2.GetSplitCurves(points);
            Polyline curves = new Polyline();
            for (int i = 0; i < curves2.Count; i++)
            {
                var pline = (Polyline)curves2[i];
                if (pline.StartPoint == points[0] && pline.EndPoint == points[1])
                {
                    curves = pline;
                }
                else if (pline.EndPoint == points[0] && pline.StartPoint == points[1])
                {
                    curves = pline;
                }
            }
            var startIndex = 0;
            var endIndex = 0;
            for (int i = 0; i < poly1.NumberOfVertices-1; i++)
            {
                var line = new Line(poly1.GetPoint3dAt(i), poly1.GetPoint3dAt(i + 1));
                if (IsPointOnLine(points[0].toPoint2d(),line))
                {
                    startIndex = i;
                }
                if (IsPointOnLine(points[1].toPoint2d(), line))
                {
                    endIndex = i;
                }
            }

            if (startIndex>endIndex)
            {
                var temp = startIndex;
                startIndex = endIndex;
                endIndex = temp;
            }
            poly1.UpgradeOpen();
            for (int j = endIndex; j> startIndex; j--)
            {
                poly1.RemoveVertexAt(j);
            }

            for (int i = 0; i < curves.NumberOfVertices; i++)
            {
                startIndex++;
                poly1.AddVertexAt(startIndex,curves.GetPoint2dAt(i),0,0,0);
            }
           
            poly1.DowngradeOpen();
         
            poly2.UpgradeOpen();
            poly2.Erase();
            poly2.DowngradeOpen();

        }
      /// <summary>
        /// 判断2d点是否在线段上
        /// </summary>
        /// <param name="p"></param>
        /// <param name="line"></param>
        /// <returns></returns>
        public bool IsPointOnLine(Point2d p, Line line)
        {
            var flag = p.GetDistanceTo(line.StartPoint.toPoint2d()) + p.GetDistanceTo(line.EndPoint.toPoint2d()) - line.Length < 0.00001;
            return flag;
        }

        // 多段线和直线求交点
        private void PolyIntersectWithLine(Polyline poly, Line line, double tol, ref Point3dCollection points)
        {
            Point2dCollection intPoints2d = new Point2dCollection();

            // 获得直线对应的几何类
            LineSegment2d geLine = new LineSegment2d(ToPoint2d(line.StartPoint), ToPoint2d(line.EndPoint));

            // 每一段分别计算交点
            Tolerance tolerance = new Tolerance(tol, tol);
            for (int i = 0; i < poly.NumberOfVertices; i++)
            {
                if (i < poly.NumberOfVertices - 1 || poly.Closed)
                {
                    SegmentType st = poly.GetSegmentType(i);
                    if (st == SegmentType.Line)
                    {
                        LineSegment2d geLineSeg = poly.GetLineSegment2dAt(i);
                        Point2d[] pts = geLineSeg.IntersectWith(geLine, tolerance);
                        if (pts != null)
                        {
                            for (int j = 0; j < pts.Length; j++)
                            {
                                if (FindPointIn(intPoints2d, pts[j], tol) < 0)
                                {
                                    intPoints2d.Add(pts[j]);
                                }
                            }
                        }
                    }
                    else if (st == SegmentType.Arc)
                    {
                        CircularArc2d geArcSeg = poly.GetArcSegment2dAt(i);
                        Point2d[] pts = geArcSeg.IntersectWith(geLine, tolerance);
                        if (pts != null)
                        {
                            for (int j = 0; j < pts.Length; j++)
                            {
                                if (FindPointIn(intPoints2d, pts[j], tol) < 0)
                                {
                                    intPoints2d.Add(pts[j]);
                                }
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < intPoints2d.Count; i++)
            {
                points.Add(ToPoint3d(intPoints2d[i]));
            }
        }

        // 点是否在集合中
        private int FindPointIn(Point2dCollection points, Point2d pt, double tol)
        {
            for (int i = 0; i < points.Count; i++)
            {
                if (Math.Abs(points[i].X - pt.X) < tol && Math.Abs(points[i].Y - pt.Y) < tol)
                {
                    return i;
                }
            }

            return -1;
        }

        // 三维点转二维点
        private static Point2d ToPoint2d( Point3d point3d)
        {
            return new Point2d(point3d.X, point3d.Y);
        }

        // 二维点转三维点
        private static Point3d ToPoint3d(Point2d point2d)
        {
            return new Point3d(point2d.X, point2d.Y, 0);
        }

        private static Point3d ToPoint3d(Point2d point2d, double elevation)
        {
            return new Point3d(point2d.X, point2d.Y, elevation);
        }
    }
}
