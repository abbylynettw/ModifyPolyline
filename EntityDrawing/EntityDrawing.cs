using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace EntityDrawing
{
    // 用于添加实体的绘制的操作
    public class EntityDrawing : DrawJig
    {
        private List<Entity> list = new List<Entity>();//表示需要进行绘制的实体
        /// <summary>
        /// 通过委托 向外显示出 需要绘制的实体集合 以及拖拽类的焦点的设置
        /// 第一点是 拖拽不动的一点
        /// 第二点是 随鼠标变化的点
        /// </summary>
        private Func<Point3d, List<Entity>> mOption = null;
        private string message = "";
        private Point3d basePoint = Point3d.Origin;
        /// <summary>
        ///  添加一些可以预览的实体 但是不能使用参照块
        ///  判断是否取消拖拽
        /// </summary>
        /// <param name="ents">  </param>
        public EntityDrawing(Point3d basePoint, Func<Point3d, List<Entity>> options, string message)
        {
            this.mOption = options;
            this.message = message;
            this.basePoint = basePoint;
            if (options == null) throw new ArgumentException("拖拽类的操作不能空！");
        }
        protected override bool WorldDraw(Autodesk.AutoCAD.GraphicsInterface.WorldDraw draw)
        {
            if (list != null)
            {
                foreach (var ent in list)
                {
                    draw.Geometry.Draw(ent);
                }
            }
            return true;
        }
        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Matrix3d mt = ed.CurrentUserCoordinateSystem;
            JigPromptPointOptions optJigPoint = new JigPromptPointOptions(message);
            optJigPoint.Cursor = CursorType.Crosshair;
            optJigPoint.UserInputControls = UserInputControls.Accept3dCoordinates | UserInputControls.NoZeroResponseAccepted | UserInputControls.NoNegativeResponseAccepted | UserInputControls.NullResponseAccepted;
            optJigPoint.BasePoint = basePoint.TransformBy(mt);//将wcs 转为ucs坐标
            optJigPoint.UseBasePoint = true;
            PromptPointResult result = prompts.AcquirePoint(optJigPoint);
            Point3d tempPt = result.Value;
            if (result.Status == PromptStatus.Cancel)
            {
                return SamplerStatus.Cancel;
            }
            Point3d temp = result.Value;
            temp = temp.TransformBy(mt.Inverse());// 将选择点 转化为当前的用户坐标

            if (basePoint != result.Value)// 如果坐标发生了变化 重新绘制当前的图形
            {
                list = mOption.Invoke(temp.TransformBy(mt));
                //basePoint = temp;
                return SamplerStatus.OK;
            }
            else
            {
                return SamplerStatus.NoChange;
            }
        }

    }
}
