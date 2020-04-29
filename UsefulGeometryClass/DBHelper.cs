using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace UsefulGeometryClass
{
    public static class DBHelper
    {
        /// <summary>
        /// 将实体添加到特定空间。
        /// </summary>
        /// <param name="ent"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public static ObjectId ToSpace(this Entity ent, Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                id = mdlSpc.AppendEntity(ent);
                trans.AddNewlyCreatedDBObject(ent, true);

                trans.Commit();
            }

            return id;
        }
        /// <summary>
        /// 调用COM的SendCommand函数
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="args">命令参数列表</param>
        public static void SendCommand(this Document doc, params string[] args)
        {
            Type AcadDocument = Type.GetTypeFromHandle(Type.GetTypeHandle(doc.GetAcadDocument()));
            try
            {
                // 通过后期绑定的方式调用SendCommand命令
                AcadDocument.InvokeMember("SendCommand", BindingFlags.InvokeMethod, null, doc.GetAcadDocument(), args);
            }
            catch // 捕获异常
            {
                return;
            }
        }
        public static Point2d toPoint2d(this Point3d p)
        {
            return new Point2d(p.X, p.Y);
        }
        /// <summary>
        /// 将实体集合添加到特定空间。
        /// </summary>
        /// <param name="ents"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public static ObjectIdCollection ToSpace(this IEnumerable<Entity> ents,
            Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var ids = new ObjectIdCollection();

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                foreach (var ent in ents)
                {
                    ids.Add(mdlSpc.AppendEntity(ent));
                    trans.AddNewlyCreatedDBObject(ent, true);
                }

                trans.Commit();
            }

            return ids;
        }

        /// <summary>
        /// 对实体集合进行遍历并操作。
        /// </summary>
        /// <param name="ents"></param>
        /// <param name="act"></param>
        public static void ForEach(this IEnumerable<Entity> ents, Action<Entity> act)
        {
            foreach (var ent in ents)
            {
                act.Invoke(ent);
            }
        }

        /// <summary>
        /// 遍历特定空间，对实体进行操作。
        /// </summary>
        /// <param name="act"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        public static void ForEach(Action<Entity> act, Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace],
                    OpenMode.ForRead) as BlockTableRecord;

                foreach (var id in mdlSpc)
                {
                    var ent = trans.GetObject(id, OpenMode.ForWrite) as Entity;
                    act.Invoke(ent);
                }

                trans.Commit();
            }
        }

        /// <summary>
        /// 将实体集合转化为块定义。
        /// </summary>
        /// <param name="ents"></param>
        /// <param name="blockName"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static ObjectId ToBlockDefinition(this IEnumerable<Entity> ents,
            string blockName, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            var blkDef = new BlockTableRecord();
            blkDef.Name = blockName;

            foreach (var ent in ents)
            {
                blkDef.AppendEntity(ent);
            }

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

                id = blkTbl.Add(blkDef);
                trans.AddNewlyCreatedDBObject(blkDef, true);

                trans.Commit();
            }

            return id;
        }

        /// <summary>
        /// 将块参照插入到特定空间。
        /// </summary>
        /// <param name="blkDefId"></param>
        /// <param name="transform"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public static ObjectId Insert(ObjectId blkDefId, Matrix3d transform,
            Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                var mdlSpc = trans.GetObject(blkTbl[space ?? BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                var blkRef = new BlockReference(Point3d.Origin, blkDefId);
                blkRef.BlockTransform = transform;

                id = mdlSpc.AppendEntity(blkRef);
                trans.AddNewlyCreatedDBObject(blkRef, true);

                // 添加属性文字。
                var blkDef = trans.GetObject(blkDefId, OpenMode.ForRead) as BlockTableRecord;
                if (blkDef.HasAttributeDefinitions)
                {
                    foreach (var subId in blkDef)
                    {
                        if (subId.ObjectClass.Equals(RXClass.GetClass(typeof(AttributeDefinition))))
                        {
                            var attrDef = trans.GetObject(subId, OpenMode.ForRead) as AttributeDefinition;
                            var attrRef = new AttributeReference();
                            attrRef.SetAttributeFromBlock(attrDef, transform);

                            blkRef.AttributeCollection.AppendAttribute(attrRef);
                        }
                    }
                }

                trans.Commit();
            }

            return id;
        }

        /// <summary>
        /// 制定块定义名称，将块参照插入到特定空间。
        /// </summary>
        /// <param name="name"></param>
        /// <param name="transform"></param>
        /// <param name="db"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public static ObjectId Insert(string name, Matrix3d transform, Database db = null, string space = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var blkTbl = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                if (!blkTbl.Has(name))
                {
                    return ObjectId.Null;
                }
                id = blkTbl[name];
            }

            return Insert(id, transform, db, space);
        }

        /// <summary>
        /// 获取符号。
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="name"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static ObjectId GetSymbol(ObjectId tblId, string name, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var tbl = trans.GetObject(tblId, OpenMode.ForRead) as SymbolTable;
                if (tbl.Has(name))
                {
                    return tbl[name];
                }
            }

            return ObjectId.Null;
        }

        /// <summary>
        /// 新建符号表记录。
        /// </summary>
        /// <param name="record"></param>
        /// <param name="tblId"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static ObjectId ToTable(this SymbolTableRecord record, ObjectId tblId,
            Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var id = ObjectId.Null;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var tbl = trans.GetObject(tblId, OpenMode.ForWrite) as SymbolTable;
                if (tbl.Has(record.Name))
                {
                    return tbl[record.Name];
                }

                tbl.Add(record);
                trans.AddNewlyCreatedDBObject(record, true);

                trans.Commit();
            }

            return id;
        }

        /// <summary>
        /// 修改符号表记录。
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="name"></param>
        /// <param name="act"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static bool ModifySymbol(ObjectId tblId, string name,
            Action<SymbolTableRecord> act, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var tbl = trans.GetObject(tblId, OpenMode.ForRead) as SymbolTable;
                if (!tbl.Has(name))
                {
                    return false;
                }

                var symbol = trans.GetObject(tbl[name], OpenMode.ForWrite) as SymbolTableRecord;
                act.Invoke(symbol);

                trans.Commit();
            }

            return true;
        }

        /// <summary>
        /// 删除符号表记录。
        /// </summary>
        /// <param name="tblId"></param>
        /// <param name="name"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static bool RemoveSymbol(ObjectId tblId, string name, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var tbl = trans.GetObject(tblId, OpenMode.ForWrite) as SymbolTable;
                if (!tbl.Has(name))
                {
                    return false;
                }

                SymbolTableRecord inUse = null;
                ObjectId inUseId = ObjectId.Null;

                if (tbl is LayerTable)
                {
                    inUseId = db.Clayer;
                }
                else if (tbl is TextStyleTable)
                {
                    inUseId = db.Textstyle;
                }
                else if (tbl is DimStyleTable)
                {
                    inUseId = db.Dimstyle;
                }
                else if (tbl is LinetypeTable)
                {
                    inUseId = db.Celtype;
                }

                if (inUseId.IsValid)
                {
                    inUse = trans.GetObject(inUseId, OpenMode.ForRead) as SymbolTableRecord;
                    if (inUse.Name.ToUpper() == name.ToUpper())
                    {
                        return false;
                    }
                }

                var record = trans.GetObject(tbl[name], OpenMode.ForWrite);
                if (record.IsErased)
                {
                    return false;
                }

                var idCol = new ObjectIdCollection() { record.ObjectId };
                db.Purge(idCol);
                if (idCol.Count == 0)
                {
                    return false;
                }

                record.Erase();
                trans.Commit();
            }

            return true;
        }

        /// <summary>
        /// 向数据库对象添加扩展数据。
        /// 可通过传递空数组实现删除。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="app"></param>
        /// <param name="datas"></param>
        /// <param name="db"></param>
        public static void AttachXData(this DBObject obj, string app,
            IEnumerable<TypedValue> datas, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            if (GetSymbol(db.RegAppTableId, app, db) == ObjectId.Null)
            {
                new RegAppTableRecord()
                {
                    Name = app,
                }.ToTable(db.RegAppTableId, db);
            }

            var rb = new ResultBuffer();
            rb.Add(new TypedValue((int)DxfCode.ExtendedDataRegAppName, app));
            foreach (var data in datas)
            {
                rb.Add(data);
            }

            obj.XData = rb;
        }

        /// <summary>
        /// 向数据库对象添加扩展数据。
        /// 可通过传递空数组实现删除。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="app"></param>
        /// <param name="datas"></param>
        /// <param name="db"></param>
        public static void AttachXData(this ObjectId objId, string app,
            IEnumerable<TypedValue> datas, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForWrite);
                obj.AttachXData(app, datas, db);

                trans.Commit();
            }
        }

        /// <summary>
        /// 获取数据库对象上的扩展数据。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="app"></param>
        /// <returns></returns>
        public static TypedValue[] GetXData(this DBObject obj, string app)
        {
            return obj.GetXDataForApplication(app)?.AsArray()?.Skip(1)?.ToArray();
        }

        /// <summary>
        /// 获取数据库对象上的扩展数据。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="app"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static TypedValue[] GetXData(this ObjectId objId, string app,
            Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForRead);
                return obj.GetXData(app);
            }
        }

        /// <summary>
        /// 设置数据库对象上的扩展数据（单个值）。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="app"></param>
        /// <param name="idx"></param>
        /// <param name="newVal"></param>
        /// <returns></returns>
        public static TypedValue? SetXData(this DBObject obj, string app, int idx,
            TypedValue newVal)
        {
            var valArr = obj.GetXDataForApplication(app)?.AsArray();
            if (valArr != null && idx + 1 < valArr.Length)
            {
                var oldVal = valArr[idx + 1];

                valArr[idx + 1] = newVal;
                obj.XData = new ResultBuffer(valArr);

                return oldVal;
            }

            return null;
        }

        /// <summary>
        /// 设置数据库对象上的扩展数据（单个值）。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="app"></param>
        /// <param name="idx"></param>
        /// <param name="newVal"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static TypedValue? SetXData(this ObjectId objId, string app, int idx,
            TypedValue newVal, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForRead);
                return obj.SetXData(app, idx, newVal);
            }
        }

        /// <summary>
        /// 向数据库对象的扩展字典设置值。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="app"></param>
        /// <param name="datas"></param>
        /// <param name="db"></param>
        public static void SetExtData(this DBObject obj, string key, DBObject data,
            Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            if (!obj.ExtensionDictionary.IsValid)
            {
                obj.CreateExtensionDictionary();
            }

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var dict = trans.GetObject(obj.ExtensionDictionary,
                    OpenMode.ForWrite) as DBDictionary;
                dict.SetAt(key, data);
            }
        }

        /// <summary>
        /// 向数据库对象的扩展字典设置值。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="app"></param>
        /// <param name="datas"></param>
        /// <param name="db"></param>
        public static void SetExtData(this ObjectId objId, string key, DBObject data,
            Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForWrite);
                obj.SetExtData(key, data, db);
            }
        }

        /// <summary>
        /// 获取数据库对象的扩展字典数据。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="key"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static ObjectId GetExtData(this DBObject obj, string key, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            if (obj.ExtensionDictionary.IsValid)
            {
                using (var trans = db.TransactionManager.StartTransaction())
                {
                    var dict = trans.GetObject(obj.ExtensionDictionary,
                        OpenMode.ForRead) as DBDictionary;
                    if (dict.Contains(key))
                    {
                        return dict.GetAt(key);
                    }
                }
            }
            return ObjectId.Null;
        }

        /// <summary>
        /// 获取数据库对象的扩展字典数据。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="key"></param>
        /// <param name="db"></param>
        /// <returns></returns>
        public static ObjectId GetExtData(this ObjectId objId, string key, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForRead);
                return obj.GetExtData(key, db);
            }
        }

        /// <summary>
        /// 修改扩展字典中的数据。
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="act"></param>
        /// <param name="key"></param>
        /// <param name="db"></param>
        public static void ModifyExtData(this DBObject obj, Action<DBObject> act,
            string key, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;
            var dataId = obj.GetExtData(key, db);

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var data = trans.GetObject(dataId, OpenMode.ForWrite);
                act.Invoke(data);
            }
        }

        /// <summary>
        /// 修改扩展字典中的数据。
        /// </summary>
        /// <param name="objId"></param>
        /// <param name="act"></param>
        /// <param name="key"></param>
        /// <param name="db"></param>
        public static void ModifyExtData(this ObjectId objId, Action<DBObject> act,
            string key, Database db = null)
        {
            db = db ?? Application.DocumentManager.MdiActiveDocument.Database;

            using (var trans = db.TransactionManager.StartTransaction())
            {
                var obj = trans.GetObject(objId, OpenMode.ForRead);
                obj.ModifyExtData(act, key, db);
            }
        }
        public static BlockTableRecord GetBlock(Document doc, Transaction transaction, string blockname)
        {

            BlockTable bt = (BlockTable)transaction.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            foreach (ObjectId blockTableRecord in bt)
            {
                BlockTableRecord btr = blockTableRecord.GetObject(OpenMode.ForRead) as BlockTableRecord;
                if (btr != null && btr.Name == blockname)
                {
                    return btr;
                }
            }
            return null;
        }
        public static bool DelBlock(Document doc, Transaction transaction, BlockTableRecord BlockTableRecord_)
        {
            try
            {
                if (BlockTableRecord_.ObjectId != ObjectId.Null)
                {
                    DBObject obj = transaction.GetObject(BlockTableRecord_.ObjectId, OpenMode.ForWrite);
                    if (obj != null)
                    {
                        obj.Erase();

                        return true;
                    }
                    return false;
                }
                return false;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception EX)
            {
                return false;
            }
        }
        public static BlockTableRecord GetBlock(Document doc, Transaction transaction, string blockName, string dwgPath)
        {
            BlockTableRecord block = GetBlock(doc, transaction, blockName);
            if (block == null)
            {
                try
                {
                    string dwgBlockName = string.Empty;
                    using (Database newDB = new Database(false, true))
                    {
                        // 读取图纸数据库
                        newDB.ReadDwgFile(dwgPath, FileOpenMode.OpenForReadAndAllShare, true, null);
                        bool isequal = false;
                        using (Transaction tr = newDB.TransactionManager.StartTransaction())
                        {
                            BlockTable bt = (BlockTable)tr.GetObject(newDB.BlockTableId, OpenMode.ForRead);
                            foreach (var oid in bt)
                            {
                                BlockTableRecord btr = oid.GetObject(OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null)
                                {
                                    if (btr.Name.Equals(blockName))
                                    {
                                        isequal = true;
                                    }
                                }
                            }
                            foreach (var oid in bt)
                            {
                                BlockTableRecord btr = oid.GetObject(OpenMode.ForRead) as BlockTableRecord;
                                if (btr != null)
                                {
                                    if (!isequal && btr.Name.IndexOf("TD_") == 0)
                                    {
                                        //根据图元图纸中的唯一块名 TD_ 找到块 
                                        dwgBlockName = btr.Name;//得到图元的块记录名
                                        break;
                                    }
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(dwgBlockName))
                        {
                            BlockTableRecord nowblock = GetBlock(doc, transaction, dwgBlockName);
                            if (nowblock == null)
                            {
                                //当前数据库中不存在要插入的块
                                ObjectId idBTR = doc.Database.Insert(blockName, newDB, false);
                                BlockTableRecord newbtr = transaction.GetObject(idBTR, OpenMode.ForRead) as BlockTableRecord;
                                nowblock = GetBlock(doc, transaction, dwgBlockName);//得到图元中的块记录
                                DelBlock(doc, transaction, newbtr);//删除新建的块记录
                            }
                            // 如果存在，直接使用
                            block = nowblock;
                        }
                        else
                        {
                            if (isequal)
                            {
                                BlockTableRecord nowblock = GetBlock(doc, transaction, blockName);
                                if (nowblock == null)
                                {
                                    ObjectId idBTR = doc.Database.Insert(blockName + Guid.NewGuid().ToString("N"), newDB, false);
                                    BlockTableRecord newbtr = transaction.GetObject(idBTR, OpenMode.ForRead) as BlockTableRecord;
                                    nowblock = GetBlock(doc, transaction, blockName);//得到图元中的块记录
                                    DelBlock(doc, transaction, newbtr);//删除新建的块记录
                                    block = nowblock;
                                }
                            }
                            else
                            {
                                ObjectId idBTR = doc.Database.Insert(blockName, newDB, false);
                                block = idBTR.GetObject(OpenMode.ForRead) as BlockTableRecord;
                            }
                        }



                    }
                }
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    throw ex;
                }
            }
            return block;
        }
        public static ObjectId InsertBlockRefByDWGFile(Document doc, Transaction transaction, string dwgPath, Point3d point, Matrix3d? mtx, string blockName = "")
        {
            ObjectId objectId = ObjectId.Null;
            blockName = string.IsNullOrEmpty(blockName) ? Guid.NewGuid().ToString("N") : Regex.Replace(blockName, @"[\u005c\u003a]", "X");
            BlockTableRecord block = GetBlock(doc, transaction, blockName, dwgPath);

            BlockTable bt = (BlockTable)transaction.GetObject(doc.Database.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)transaction.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
            using (BlockReference bref = new BlockReference(point, block.ObjectId))
            {
                if (mtx != null)
                {
                    bref.TransformBy(mtx.Value);
                }
                objectId = btr.AppendEntity(bref);
                transaction.AddNewlyCreatedDBObject(bref, true);
                return objectId;
            }

        }
    }
}
