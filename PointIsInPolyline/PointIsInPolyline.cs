using System;
using System.Diagnostics;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace PointIsInPolyline
{
    public class PointIsInPolyline
    {
        // 测试许多点是否在多段线内部
        [CommandMethod("TestPtInPoly")]
        public void TestPtInPoly()
        {
            // 随机点测试的数量
            int count = 100000;
            //if (GetInputInteger("\n输入需要测试的点数量:", ref count))
            {
                // 提示用户选择一条多段线
                ObjectId polyId = new ObjectId();
                Point3d pt = new Point3d();
                if (PromptSelectEntity("\n选择需要进行测试的多段线:", ref polyId, ref pt))
                {
                    Database db = HostApplicationServices.WorkingDatabase;
                    using (Transaction trans = db.TransactionManager.StartTransaction())
                    {
                        Polyline poly = trans.GetObject(polyId, OpenMode.ForWrite) as Polyline;
                        if (poly != null)
                        {
                            BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                            // 在多段线包围框范围内随机生成点，测试点和多段线的位置关系
                            Extents3d ext = poly.GeometricExtents;
                            double margin = 10;
                            double xmin = ext.MinPoint.X - margin;
                            double ymin = ext.MinPoint.Y - margin;
                            double xSpan = ext.MaxPoint.X - ext.MinPoint.X + 2 * margin;
                            double ySpan = ext.MaxPoint.Y - ext.MinPoint.Y + 2 * margin;

                            for (int i = 0; i < count; i++)
                            {
                                Random rand = new Random(GetRandomSeed());
                                int xRand = rand.Next();
                                int yRand = rand.Next();
                                Point2d ptTest = new Point2d(xmin + (Convert.ToDouble(xRand) / int.MaxValue) * xSpan,
                                    ymin + (Convert.ToDouble(yRand) / int.MaxValue) * ySpan);
                                int relation = PtRelationToPoly(poly, ptTest, 1.0E-4);

                                DBPoint dbPoint = new DBPoint(ToPoint3d(ptTest));
                                switch (relation)
                                {
                                    case -1:
                                        dbPoint.ColorIndex = 1;
                                        break;
                                    case 0:
                                        dbPoint.ColorIndex = 5;
                                        break;
                                    case 1:
                                        dbPoint.ColorIndex = 6;
                                        break;
                                    default:
                                        break;
                                }
                                btr.AppendEntity(dbPoint);
                                trans.AddNewlyCreatedDBObject(dbPoint, true);
                            }

                            trans.Commit();
                        }
                        else
                        {
                            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
                            ed.WriteMessage("\n选择的实体不是多段线.");
                        }
                    }
                }
            }
        }

        // 获得随机数种子
        static int GetRandomSeed()
        {
            byte[] bytes = new byte[4];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(bytes);

            return BitConverter.ToInt32(bytes, 0);
        }

        // 判断点和多段线的位置关系
        // 返回值：-1表示在多段线外部，0表示在多段线上，1表示在多段线内部
        private int PtRelationToPoly(Polyline pPoly, Point2d pt, double tol)
        {
            Debug.Assert(pPoly != null);

            // 1.如果点到多段线的最近点和给定的点重合，表示点在多段线上
            Point3d closestPoint = pPoly.GetClosestPointTo(ToPoint3d(pt, pPoly.Elevation), false);		// 多段线上与给定点距离最近的点	
            if (Math.Abs(closestPoint.X - pt.X) < tol && Math.Abs(closestPoint.Y - pt.Y) < tol)			// 点在多段线上
            {
                return 0;
            }

            // 2.第一个射线的方向是从最近点到当前点，起点是当前点
            Ray pRay = new Ray();
            pRay.BasePoint = new Point3d(pt.X, pt.Y, pPoly.Elevation);
            // 射线的起点是pt，方向为从最近点到pt，如果反向做判断，则最近点距离pt太近的时候，最近点也会被作为一个交点（这个交点不太容易被排除掉）
            // 此外，这样的射线方向很容易判断出点不在内部的情况
            Vector3d vec = new Vector3d(-(closestPoint.X - pt.X), -(closestPoint.Y - pt.Y), 0);
            pRay.UnitDir = vec;

            // 3.射线与多段线计算交点
            Point3dCollection intPoints = new Point3dCollection();
            
            pPoly.IntersectWith(pRay, Intersect.OnBothOperands, intPoints, 0, 0);
            // IntersectWith函数经常会得到很近的交点，这些点必须进行过滤
            FilterEqualPoints(intPoints, 1.0E-4);

            // 4.判断点和多段线的位置关系
        RETRY:
            Point3d[] pts = new Point3d[10];		//////////////////////////////////////////////////////////////////////////
            if (intPoints.Count > 0)
            {
                pts[0] = intPoints[0];
            }
            if (intPoints.Count > 1)
            {
                pts[1] = intPoints[1];
            }
            if (intPoints.Count > 2)
            {
                pts[2] = intPoints[2];
            }
            if (intPoints.Count > 3)
            {
                pts[3] = intPoints[3];
            }
            // 4.1 如果射线和多段线没有交点，表示点在多段线的外部
            if (intPoints.Count == 0)
            {
                pRay.Dispose();
                return -1;
            }
            else
            {
                // 3.1 过滤掉由于射线被反向延长带来的影响
                FilterEqualPoints(intPoints, ToPoint2d(closestPoint), 1.0E-4);		// 2008-0907修订记录：当pt距离最近点比较近的时候，最近点竟然被作为一个交点！
                // 3.2 如果某个交点与最近点在给定点的同一方向，要去掉这个点（这个点明显不是交点，还是由于intersectwith函数的Bug）	
                for (int i = intPoints.Count - 1; i >= 0; i--)
                {
                    if ((intPoints[i].X - pt.X) * (closestPoint.X - pt.X) >= 0 &&
                        (intPoints[i].Y - pt.Y) * (closestPoint.Y - pt.Y) >= 0)
                    {
                        intPoints.RemoveAt(i);
                    }
                }

                int count = intPoints.Count;
                for (int i = 0; i < intPoints.Count; i++)
                {
                    if (PointIsPolyVert(pPoly, ToPoint2d(intPoints[i]), 1.0E-4))		// 只要有交点是多段线的顶点就重新进行判断
                    {
                        // 处理给定点很靠近多段线顶点的情况(如果与顶点距离很近，就认为这个点在多段线上，因为这种情况没有什么好的判断方法)
                        if (PointIsPolyVert(pPoly, new Point2d(pt.X, pt.Y), 1.0E-4))
                        {
                            return 0;
                        }

                        // 将射线旋转一个极小的角度(2度)再次判断（假定这样不会再通过上次判断到的顶点）
                        vec = vec.RotateBy(0.035, Vector3d.ZAxis);
                        pRay.UnitDir = vec;
                        intPoints.Clear();
                        pPoly.IntersectWith(pRay, Intersect.OnBothOperands, intPoints, 0, 0);
                        goto RETRY;		// 继续判断结果
                    }
                }

                pRay.Dispose();

                if (count % 2 == 0)
                {
                    return -1;
                }
                else
                {
                    return 1;
                }
            }
        }

        // 三维点转二维点
        private static Point2d ToPoint2d(Point3d point3d)
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

        // 从点数组中删除与给定点平面位置重合的点
        // tol: 判断点重合时的精度（两点之间的距离小于tol认为这两个点重合）
        static void FilterEqualPoints(Point3dCollection points, Point2d pt, double tol)
        {
            Point3dCollection tempPoints = new Point3dCollection();
            for (int i = 0; i < points.Count; i++)
            {
                if (ToPoint2d(points[i]).GetDistanceTo(pt) > tol)
                {
                    tempPoints.Add(points[i]);
                }
            }

            points = tempPoints;
        }

        // 从点数组中删除与其他点重合的点
        static void FilterEqualPoints(Point3dCollection points, double tol)
        {
            for (int i = points.Count - 1; i > 0; i--)
            {
                for (int j = 0; j < i; j++)
                {
                    if (IsEqual(points[i].X, points[j].X, tol) && IsEqual(points[i].Y, points[j].Y, tol))
                    {
                        points.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        // 点是否是多段线的顶点
        static bool PointIsPolyVert(Polyline pPoly, Point2d pt, double tol)
        {
            for (int i = 0; i < pPoly.NumberOfVertices; i++)
            {
                Point3d vert = pPoly.GetPoint3dAt(i);

                if (IsEqual(ToPoint2d(vert), pt, tol))
                {
                    return true;
                }
            }

            return false;
        }

        // 二维点是否相同
        static bool IsEqual(Point2d firstPoint, Point2d secondPoint, double tol)
        {
            return (Math.Abs(firstPoint.X - secondPoint.X) < tol && Math.Abs(firstPoint.Y - secondPoint.Y) < tol);
        }

        // 两个实数是否相等
        static bool IsEqual(double a, double b, double tol)
        {
            return (Math.Abs(a - b) < tol);
        }

        // 提示用户选择一个实体
        public static bool PromptSelectEntity(string prompt, ref ObjectId entId, ref Point3d pt)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptEntityOptions peo = new PromptEntityOptions(prompt);
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {
                pt = per.PickedPoint;
                entId = per.ObjectId;
                return true;
            }
            else
            {
                return false;
            }
        }

        // 提示用户输入整数
        public static bool GetInputInteger(string prompt, ref int val)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptIntegerOptions pio = new PromptIntegerOptions(prompt);
            pio.LowerLimit = 1;
            pio.UpperLimit = 1000000;
            PromptIntegerResult pir = ed.GetInteger(pio);
            if (pir.Status == PromptStatus.OK)
            {
                val = pir.Value;
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
