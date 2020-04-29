using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace TranlateCoordinate
{
    public class TranlateCoordinate 
    {
        // 绘制矩形管道
        [CommandMethod("DrawRectPipe")]
        public void DrawRectPipe()
        {
            // 输入起点、终点
            Point3d startPoint = new Point3d();
            Point3d endPoint = new Point3d();
            if (GetPoint("\n输入起点:", out startPoint) && GetPoint("\n输入终点:", startPoint, out endPoint))
            {
                // 绘制管道
                DrawPipe(startPoint, endPoint, 100, 70);
            }            
        }

        // 绘制管道
        private void DrawPipe(Point3d startPoint, Point3d endPoint, double width, double height)
        {
            // 获得变换矩阵
            Vector3d inVector = endPoint - startPoint;      // 入口向量
            Vector3d normal = GetNormalByInVector(inVector);       // 法向量
            Matrix3d mat = GetTranslateMatrix(startPoint, inVector, normal);

            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {                
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                
                // 顶面
                double z = 0.5 * height;
                double length = startPoint.DistanceTo(endPoint);
                Face fTop = new Face(new Point3d(0, -0.5 * width, z), new Point3d(length, -0.5 * width, z), new Point3d(length, 0.5 * width, z), 
                    new Point3d(0, 0.5 * width, z), true, true, true, true);
                fTop.TransformBy(mat);
                btr.AppendEntity(fTop);
                trans.AddNewlyCreatedDBObject(fTop, true);

                // 底面
                z = -0.5 * height;
                Face fBottom = new Face(new Point3d(0, -0.5 * width, z), new Point3d(length, -0.5 * width, z), new Point3d(length, 0.5 * width, z),
                    new Point3d(0, 0.5 * width, z), true, true, true, true);
                fBottom.TransformBy(mat);
                btr.AppendEntity(fBottom);
                trans.AddNewlyCreatedDBObject(fBottom, true);

                // 左侧面
                double y = 0.5 * width;
                Face fLeftSide = new Face(new Point3d(0, y, 0.5 * height), new Point3d(length, y, 0.5 * height), new Point3d(length, y, -0.5 * height),
                    new Point3d(0, y, -0.5 * height), true, true, true, true);
                fLeftSide.TransformBy(mat);
                btr.AppendEntity(fLeftSide);
                trans.AddNewlyCreatedDBObject(fLeftSide, true);

                // 左侧面
                y = -0.5 * width;
                Face fRightSide = new Face(new Point3d(0, y, 0.5 * height), new Point3d(length, y, 0.5 * height), new Point3d(length, y, -0.5 * height),
                    new Point3d(0, y, -0.5 * height), true, true, true, true);
                fRightSide.TransformBy(mat);
                btr.AppendEntity(fRightSide);
                trans.AddNewlyCreatedDBObject(fRightSide, true);

                trans.Commit();
            }
        }

        // 根据入口向量、法向量获得变换矩阵
        Matrix3d GetTranslateMatrix(Point3d inPoint, Vector3d inVector, Vector3d normal)
        {
            Vector3d xAxis = inVector;
            xAxis = xAxis.GetNormal();
            Vector3d zAxis = normal;
            zAxis = zAxis.GetNormal();
            Vector3d yAxis = new Vector3d(xAxis.X, xAxis.Y, xAxis.Z);
            yAxis = yAxis.RotateBy(Math.PI * 0.5, zAxis);

            return Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, inPoint, xAxis, yAxis, zAxis);
        }

        // 提示用户拾取点
        public bool GetPoint(string prompt, out Point3d pt)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointResult ppr = ed.GetPoint(prompt);
            if (ppr.Status == PromptStatus.OK)
            {
                pt = ppr.Value;

                // 变换到世界坐标系
                Matrix3d mat = ed.CurrentUserCoordinateSystem;
                pt.TransformBy(mat);

                return true;
            }
            else
            {
                pt = new Point3d();
                return false;
            }
        }

        public bool GetPoint(string prompt, Point3d basePoint, out Point3d pt)
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            PromptPointOptions ppo = new PromptPointOptions(prompt);
            ppo.BasePoint = basePoint;
            ppo.UseBasePoint = true;
            PromptPointResult ppr = ed.GetPoint(ppo);
            if (ppr.Status == PromptStatus.OK)
            {
                pt = ppr.Value;

                // 变换到世界坐标系
                Matrix3d mat = ed.CurrentUserCoordinateSystem;
                pt.TransformBy(mat);

                return true;
            }
            else
            {
                pt = new Point3d();
                return false;
            }
        }

        // 根据用户指定的入口点向量计算法向量
        private Vector3d GetNormalByInVector(Vector3d inVector)
        {
            double tol = 1.0E-7;
            if (Math.Abs(inVector.X) < tol && Math.Abs(inVector.Y) < tol)
            {
                if (inVector.Z >= 0)
                {
                    return new Vector3d(-1, 0, 0);
                }
                else
                {
                    return Vector3d.XAxis;
                }
            }
            else
            {
                Vector2d yAxis2d = new Vector2d(inVector.X, inVector.Y);
                yAxis2d = yAxis2d.RotateBy(Math.PI * 0.5);
                Vector3d yAxis = new Vector3d(yAxis2d.X, yAxis2d.Y, 0);
                Vector3d normal = yAxis;
                normal = normal.RotateBy(Math.PI * 0.5, inVector);
                return normal;
            }
        }
    }
}
