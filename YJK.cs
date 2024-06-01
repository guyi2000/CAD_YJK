using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

[assembly: ExtensionApplication(typeof(CAD_YJK.YJK))]

/// <summary>
/// 主命名空间，所有类均在此定义
/// </summary>
namespace CAD_YJK 
{
    /// <summary>
    /// 主程序类，CAD所有接口在此处定义
    /// </summary>
    public class YJK : IExtensionApplication
    {
        List<double[]> beamRange = null;                                // 梁范围，每个元素是一个数组，每个数组由两个数组成
        List<double[]> beamLineRange = null;                            // 梁内钢筋范围，每个元素是一个数组，每个数组由两个数组成
        List<double[]> profileRange = null;                             // 截面Y范围，每个元素是一个数组，每个数组由两个数组成
        List<List<double[]>> profileSplits = null;                      // 截面X范围，每个元素代表一个梁的截面（列表），
                                                                        // 其元素是一个数组，每个数组由两个数组成
        List<List<DBText>> profileSectionText = null;                   // 获取所有截面标注信息

        List<List<AlignedDimension>> beamDimensions = null;             // 按梁分类钢筋尺寸标注信息

        List<List<BoundingPolyline>> beamLineVertical = null;           // 梁线（竖线），计算锚固长度的边线

        List<BoundingPolyline> steelBarTable = null;                    // 钢筋表，所有水平钢筋组成的表

        Dictionary<int, List<int>> steelBarTableBeam = null;            // 梁-钢筋表
        Dictionary<int, int> steelBarTableBeamReverse = null;           // 反查钢筋梁表

        Dictionary<int, Dictionary<int, Dictionary<double, List<BoundingPolyline>>>> 
            steelBarBeamProfileRows = null;                             // 梁-截面-行-钢筋表
        Dictionary<int, Dictionary<int, Dictionary<string, string>>>
            steelBarBeamProfileAnno = null;                             // 梁-截面-行-标注表

        Dictionary<double, List<int>> steelBarTableStart = null;        // 起始锚固长度钢筋表
        Dictionary<int, double> steelBarTableStartReverse = null;       // 反查起始锚固长度表

        Dictionary<double, List<int>> steelBarTableEnd = null;          // 末尾锚固长度钢筋表
        Dictionary<int, double> steelBarTableEndReverse = null;         // 反查末尾锚固长度表

        /// <summary>
        /// 初始化命令，CAD打开时执行，本程序无需进行初始化
        /// </summary>
        public void Initialize() { }

        /// <summary>
        /// 析构命令，CAD关闭前执行，本程序无需析构
        /// </summary>
        public void Terminate() { }

        /// <summary>
        /// 显示单根钢筋参数
        /// </summary>
        [CommandMethod("ShowSteel")]
        public void SelectOneSteel()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            
            if (steelBarTable == null)    // 若无钢筋表数据，则要求初始化
            {
                ed.WriteMessage("\n请先运行初始化命令 I/YJK \n");
                return;
            }

            PromptEntityResult entityResult = ed.GetEntity(
                new PromptEntityOptions("选择钢筋")
            );

            Database dwg = ed.Document.Database;
            Transaction trans = dwg.TransactionManager.StartTransaction();

            try
            {
                if (entityResult.Status == PromptStatus.OK)
                {
                    Entity e = trans.GetObject(entityResult.ObjectId, OpenMode.ForRead) as Entity;
                    if (e == null) 
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "实体未找到");

                    BoundingPolyline bp = steelBarTable.Find(
                        (BoundingPolyline b) => b == e
                    );

                    if (bp == null)
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "实体不在钢筋表中");

                    ed.WriteMessage(
                        "\n钢筋数量 {0}, 钢筋直径 {1}", 
                        bp.num, bp.diameter
                    );
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

        } // public void SelectOneSteel()

        /// <summary>
        /// 延长所有锚固长度相同的钢筋，并修改尺寸标注
        /// </summary>
        [CommandMethod("H")]
        public void ExtendAllSameSteelBar()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            if (steelBarTable == null ||
                steelBarTableStart == null || 
                steelBarTableEnd == null || 
                steelBarTableStartReverse == null || 
                steelBarTableEndReverse == null)    // 若无钢筋表数据，则要求初始化
            {
                ed.WriteMessage("\n请先运行初始化命令 I/YJK \n");
                return;
            }

            PromptEntityResult entityResult = ed.GetEntity(
                new PromptEntityOptions("选择钢筋")
            );

            Database dwg = ed.Document.Database;
            Transaction trans = dwg.TransactionManager.StartTransaction();

            double startLength = 0;                 // 起始锚固长度
            double endLength = 0;                   // 末尾锚固长度
            List<int> sameStartSteelBar = null;     // 相同起始锚固长度钢筋
            List<int> sameEndSteelBar = null;       // 相同结尾锚固长度钢筋

            try
            {
                if (entityResult.Status == PromptStatus.OK)
                {
                    // 初始化beamDimensions
                    if (beamDimensions == null)
                    {
                        beamDimensions = new List<List<AlignedDimension>>();
                        for (int i = 0; i < beamRange.Count; i++)
                        {
                            beamDimensions.Add(new List<AlignedDimension>());
                        }

                        PromptSelectionResult allDimesions = ed.SelectAll(
                            new SelectionFilter(
                                new TypedValue[] {
                                    new TypedValue((int)DxfCode.LayerName, Utils.DIM_BEAM_NAME)
                                }
                            )
                        );

                        // 按梁号分类标注信息
                        foreach (ObjectId item in allDimesions.Value.GetObjectIds())
                        {
                            AlignedDimension ad = trans.GetObject(item, OpenMode.ForRead) as AlignedDimension;
                            if (ad == null) continue;

                            int beamIndex = beamRange.FindIndex(
                                (double[] t) => t[0] <= ad.XLine1Point.Y && ad.XLine1Point.Y <= t[1]
                            );
                            if (beamIndex == -1) continue;

                            beamDimensions[beamIndex].Add(ad);
                        }
                    }

                    Entity e = trans.GetObject(entityResult.ObjectId, OpenMode.ForRead) as Entity;
                    if (e == null) 
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "实体未找到");

                    // 获取选择钢筋的锚固长度以及相同长度的锚固钢筋
                    int steelIndex = steelBarTable.FindIndex(
                        (BoundingPolyline b) => b == e
                    );
                    startLength = steelBarTableStartReverse[steelIndex];
                    endLength = steelBarTableEndReverse[steelIndex];
                    sameStartSteelBar = new List<int>(steelBarTableStart[startLength]);
                    sameEndSteelBar = new List<int>(steelBarTableEnd[endLength]);

                    ed.WriteMessage(
                        "\n起始锚固长度为 " +
                        startLength.ToString() + 
                        " ，末尾锚固长度为 " + 
                        endLength.ToString() + 
                        "\n有相同起始锚固长度的钢筋 " +
                        sameStartSteelBar.Count.ToString() +
                        " 条，有相同末尾锚固长度的钢筋 " +
                        sameEndSteelBar.Count.ToString() +
                        " 条\n"
                    );

                }  // if (entityResult.Status == PromptStatus.OK)

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

            // 修改起始锚固长度
            trans = dwg.TransactionManager.StartTransaction(); 
            try
            {
                PromptResult startModifyResult = ed.GetKeywords(
                    new PromptKeywordOptions("是否修改起始锚固长度 [Y/N] : ", "Y N")
                );
                if (startModifyResult.Status == PromptStatus.OK &&
                    "Y".Equals(startModifyResult.StringResult))
                {
                    PromptDoubleResult modifyStartMLength = ed.GetDouble(
                        new PromptDoubleOptions("请输入修改后的起始锚固长度")
                    );
                    foreach (var lineIndex in sameStartSteelBar)
                    {
                        steelBarTable[lineIndex].GetLine().Highlight();     // 高亮显示修改线
                    }
                    PromptResult isModify = ed.GetKeywords(
                        new PromptKeywordOptions("高亮显示的线将被修改，是否确认 [Y/N] : ", "Y N")
                    );
                    foreach (var lineIndex in sameStartSteelBar)
                    {
                        steelBarTable[lineIndex].GetLine().Unhighlight();   // 取消高亮
                    }
                    if (isModify.Status == PromptStatus.OK &&
                        "Y".Equals(isModify.StringResult))
                    {
                        foreach (var lineIndex in sameStartSteelBar)
                        {
                            double extendLength = modifyStartMLength.Value - startLength;

                            Entity e = trans.GetObject(
                                steelBarTable[lineIndex].GetLine().ObjectId, 
                                OpenMode.ForWrite
                            ) as Entity;    // 以写方式打开实体
                            if (e == null) continue;

                            if (steelBarTable[lineIndex].isPolyline)
                            {
                                // 如果实体是多段线，则按如下方式延长
                                Polyline pl = e as Polyline;
                                if (pl == null) continue;

                                // 查找有无尺寸标注端点在钢筋端点处，如有，则与钢筋端点同时移动
                                Point3d previousPoint = steelBarTable[lineIndex].GetStartPoint();
                                foreach(AlignedDimension a in beamDimensions[steelBarTableBeamReverse[lineIndex]])
                                {
                                    if (a.XLine1Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine1Point -= new Vector3d(extendLength, 0, 0);
                                    }
                                    if (a.XLine2Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine2Point -= new Vector3d(extendLength, 0, 0);
                                    }
                                }

                                // 向左平移所有索引在指定线段索引前的端点
                                for (int i = 0; i <= steelBarTable[lineIndex].segmentIndex; i++)
                                {
                                    Point2d newPoint = pl.GetPoint2dAt(i) - new Vector2d(extendLength, 0);
                                    pl.SetPointAt(i, newPoint);
                                }
                                steelBarTable[lineIndex] = new BoundingPolyline(pl);
                            }
                            else
                            {
                                // 如果实体是直线，则按如下方式延长
                                Line l = e as Line;
                                if (l == null) continue;

                                // 查找有无尺寸标注端点在钢筋端点处，如有，则与钢筋端点同时移动
                                Point3d previousPoint = steelBarTable[lineIndex].GetStartPoint();
                                foreach (AlignedDimension a in beamDimensions[steelBarTableBeamReverse[lineIndex]])
                                {
                                    if (a.XLine1Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine1Point -= new Vector3d(extendLength, 0, 0);
                                    }
                                    if (a.XLine2Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine2Point -= new Vector3d(extendLength, 0, 0);
                                    }
                                }

                                l.StartPoint -= new Vector3d(extendLength, 0, 0);
                                steelBarTable[lineIndex] = new BoundingPolyline(l);
                            }

                            // 修改实体关联的钢筋表相关参数
                            steelBarTableStart[startLength].Remove(lineIndex);
                            if (steelBarTableStart.ContainsKey(Math.Round(modifyStartMLength.Value, Utils.FIXED_NUM)))
                            {
                                steelBarTableStart[Math.Round(modifyStartMLength.Value, Utils.FIXED_NUM)].Add(lineIndex);
                            }
                            else
                            {
                                steelBarTableStart.Add(
                                    Math.Round(modifyStartMLength.Value, Utils.FIXED_NUM),
                                    new List<int> { lineIndex }
                                );
                            }
                            steelBarTableStartReverse[lineIndex] = Math.Round(modifyStartMLength.Value, Utils.FIXED_NUM);
                        }
                        trans.Commit();
                        ed.WriteMessage("\n起始锚固长度修改完成");

                    }   // if (isModify.Status == PromptStatus.OK && "Y".Equals(isModify.StringResult))

                }   // if (startModifyResult.Status == PromptStatus.OK && "Y".Equals(startModifyResult.StringResult))

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

            // 修改末尾锚固长度
            trans = dwg.TransactionManager.StartTransaction();
            try
            {
                PromptResult endModifyResult = ed.GetKeywords(
                        new PromptKeywordOptions("是否修改末尾锚固长度 [Y/N] : ", "Y N")
                    );
                if (endModifyResult.Status == PromptStatus.OK &&
                    "Y".Equals(endModifyResult.StringResult))
                {
                    PromptDoubleResult modifyEndMLength = ed.GetDouble(
                        new PromptDoubleOptions("请输入修改后末尾锚固长度")
                    );
                    foreach (var lineIndex in sameEndSteelBar)
                    {
                        steelBarTable[lineIndex].GetLine().Highlight();         // 高亮显示修改线
                    }
                    PromptResult isModify = ed.GetKeywords(
                        new PromptKeywordOptions("高亮显示的线将被修改，是否确认 [Y/N] : ", "Y N")
                    );
                    foreach (var lineIndex in sameEndSteelBar)
                    {
                        steelBarTable[lineIndex].GetLine().Unhighlight();       // 取消高亮
                    }
                    if (isModify.Status == PromptStatus.OK &&
                        "Y".Equals(isModify.StringResult))
                    {
                        foreach (var lineIndex in sameEndSteelBar)
                        {
                            double extendLength = modifyEndMLength.Value - endLength;

                            Entity e = trans.GetObject(
                                steelBarTable[lineIndex].GetLine().ObjectId,
                                OpenMode.ForWrite
                            ) as Entity;            // 以写方式打开实体
                            if (e == null) continue;

                            if (steelBarTable[lineIndex].isPolyline)
                            {
                                // 如果实体是多段线，则按如下方式延长
                                Polyline pl = e as Polyline;
                                if (pl == null) continue;

                                // 查找有无尺寸标注端点在钢筋端点处，如有，则与钢筋端点同时移动
                                Point3d previousPoint = steelBarTable[lineIndex].GetEndPoint();
                                foreach (AlignedDimension a in beamDimensions[steelBarTableBeamReverse[lineIndex]])
                                {
                                    if (a.XLine1Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine1Point += new Vector3d(extendLength, 0, 0);
                                    }
                                    if (a.XLine2Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine2Point += new Vector3d(extendLength, 0, 0);
                                    }
                                }

                                // 向右平移所有索引在指定线段索引后的端点
                                for (int i = steelBarTable[lineIndex].segmentIndex + 1; i < pl.NumberOfVertices; i++)
                                {
                                    Point2d newPoint = pl.GetPoint2dAt(i) + new Vector2d(extendLength, 0);
                                    pl.SetPointAt(i, newPoint);
                                }
                                steelBarTable[lineIndex] = new BoundingPolyline(pl);
                            }
                            else
                            {
                                // 如果实体是直线，则按如下方式延长
                                Line l = e as Line;
                                if (l == null) continue;

                                // 查找有无尺寸标注端点在钢筋端点处，如有，则与钢筋端点同时移动
                                Point3d previousPoint = steelBarTable[lineIndex].GetEndPoint();
                                foreach (AlignedDimension a in beamDimensions[steelBarTableBeamReverse[lineIndex]])
                                {
                                    if (a.XLine1Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine1Point += new Vector3d(extendLength, 0, 0);
                                    }
                                    if (a.XLine2Point == previousPoint)
                                    {
                                        AlignedDimension aw = trans.GetObject(a.ObjectId, OpenMode.ForWrite) as AlignedDimension;
                                        aw.XLine2Point += new Vector3d(extendLength, 0, 0);
                                    }
                                }

                                l.EndPoint += new Vector3d(extendLength, 0, 0);
                                steelBarTable[lineIndex] = new BoundingPolyline(l);
                            }

                            // 修改实体关联的钢筋表相关参数
                            steelBarTableEnd[endLength].Remove(lineIndex);
                            if (steelBarTableEnd.ContainsKey(Math.Round(modifyEndMLength.Value, Utils.FIXED_NUM)))
                            {
                                steelBarTableEnd[Math.Round(modifyEndMLength.Value, Utils.FIXED_NUM)].Add(lineIndex);
                            }
                            else
                            {
                                steelBarTableEnd.Add(
                                    Math.Round(modifyEndMLength.Value, Utils.FIXED_NUM),
                                    new List<int> { lineIndex }
                                );
                            }
                            steelBarTableEndReverse[lineIndex] = Math.Round(modifyEndMLength.Value, Utils.FIXED_NUM);
                        }
                        trans.Commit();
                        ed.WriteMessage("\n末尾锚固长度修改完成");

                    }   // if (isModify.Status == PromptStatus.OK && "Y".Equals(isModify.StringResult))

                }   // if (endModifyResult.Status == PromptStatus.OK && "Y".Equals(endModifyResult.StringResult))

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

        }   // public void ExtendAllSameSteelBar()

        /// <summary>
        /// 计算dwg文件中所有钢筋锚固长度
        /// </summary>
        [CommandMethod("I")]
        public void CalculateAllSteelBar()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database dwg = ed.Document.Database;

            // 计算水平、垂直线，并将结果保存在steelBarTable
            // steelBarTableStart, steelBarTableStartReverse
            // steelBarTableEnd, steelBarTableEndReverse中
            Transaction trans = dwg.TransactionManager.StartTransaction();
            try
            {
                PromptSelectionResult selection = ed.SelectAll();
                if (selection.Status == PromptStatus.OK)
                {
                    List<BoundingPolyline> verticalLines = new List<BoundingPolyline>();
                    List<BoundingPolyline> horizontalLines = new List<BoundingPolyline>();

                    foreach (ObjectId item in selection.Value.GetObjectIds())
                    {
                        Entity e = trans.GetObject(item, OpenMode.ForRead) as Entity;
                        if (e == null) continue;

                        Line l = e as Line;
                        if (l != null)
                        {
                            if(Math.Abs(l.Delta.X) <= Utils.DOUBLE_EPS)         // 为垂直线
                            {
                                verticalLines.Add(new BoundingPolyline(l));
                            } else if(Math.Abs(l.Delta.Y) <= Utils.DOUBLE_EPS)
                            {
                                horizontalLines.Add(new BoundingPolyline(l));   // 为水平线
                            }
                        }
                        else
                        {
                            Polyline pl = e as Polyline;
                            if (pl != null)
                            {
                                BoundingPolyline bp = new BoundingPolyline(pl);
                                Vector3d deltaVector = bp.ls.EndPoint - bp.ls.StartPoint;
                                if (Math.Abs(deltaVector.X) <= Utils.DOUBLE_EPS)            // 为垂直线
                                {
                                    verticalLines.Add(bp);
                                }
                                else if (Math.Abs(deltaVector.Y) <= Utils.DOUBLE_EPS)       // 为水平线
                                {
                                    horizontalLines.Add(bp);
                                }

                            } // if (pl != null)

                        } // else 

                    } // foreach (ObjectId item in selection.Value.GetObjectIds())

                    // 初始化各类钢筋表
                    steelBarTable = new List<BoundingPolyline>();
                    steelBarTableStart = new Dictionary<double, List<int>>();
                    steelBarTableEnd = new Dictionary<double, List<int>>();
                    steelBarTableStartReverse = new Dictionary<int, double>();
                    steelBarTableEndReverse = new Dictionary<int, double>();

                    foreach (BoundingPolyline l1 in horizontalLines)
                    {
                        //计算锚固长度
                        double startMLength = double.PositiveInfinity;
                        double endMLength = double.PositiveInfinity;
                        foreach (BoundingPolyline l2 in verticalLines)
                        {
                            Point3dCollection pTemp = new Point3dCollection();
                            l1.GetLine().IntersectWith(
                                l2.GetLine(), Intersect.OnBothOperands, 
                                pTemp, IntPtr.Zero, IntPtr.Zero
                            );
                            if (pTemp.Count == 1)       // 有且只有一个交点
                            {
                                startMLength = Math.Min(
                                    pTemp[0].DistanceTo(l1.GetStartPoint()),
                                    startMLength
                                );
                                endMLength = Math.Min(
                                    pTemp[0].DistanceTo(l1.GetEndPoint()),
                                    endMLength
                                );
                            }

                        }   // foreach (BoundingPolyline l2 in verticalLines)

                        // 将钢筋添加进钢筋表
                        int steelBarIndex = steelBarTable.Count;
                        steelBarTable.Add(l1);

                        // 添加进起始锚固长度表
                        if (steelBarTableStart.ContainsKey(Math.Round(startMLength, Utils.FIXED_NUM)))
                            steelBarTableStart[Math.Round(startMLength, Utils.FIXED_NUM)].Add(steelBarIndex);
                        else
                            steelBarTableStart.Add(
                                Math.Round(startMLength, Utils.FIXED_NUM),
                                new List<int> { steelBarIndex }
                            );

                        // 添加进末尾锚固长度表
                        if (steelBarTableEnd.ContainsKey(Math.Round(endMLength, Utils.FIXED_NUM)))
                            steelBarTableEnd[Math.Round(endMLength, Utils.FIXED_NUM)].Add(steelBarIndex);
                        else
                            steelBarTableEnd.Add(
                                Math.Round(endMLength, Utils.FIXED_NUM),
                                new List<int> { steelBarIndex }
                            );

                        // 添加反查表
                        steelBarTableStartReverse.Add(
                            steelBarIndex,
                            Math.Round(startMLength, Utils.FIXED_NUM)
                        );
                        steelBarTableEndReverse.Add(
                            steelBarIndex,
                            Math.Round(endMLength, Utils.FIXED_NUM)
                        );

                    } // foreach (BoundingPolyline l1 in horizontalLines)

                    ed.WriteMessage("\nInitialize success!");

                } // if (selection.Status == PromptStatus.OK)

            } // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

        } // public void CalculateAllSteelBar()

        /// <summary>
        /// 处理盈建科软件导出的DWG图纸
        /// </summary>
        [CommandMethod("YJK")]
        public void YJKParse()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database dwg = ed.Document.Database;

            // 选择所有纵筋、箍筋将直线改为多段线并调整全局宽度为25
            Transaction trans = dwg.TransactionManager.StartTransaction();

            try
            {
                PromptSelectionResult selection = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.Operator, "<OR"),
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_STEEL_NAME),
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_GSTEEL_NAME),
                            new TypedValue((int)DxfCode.Operator, "OR>")
                        }
                    )
                );      // 过滤所有图层名为梁截面纵筋或梁截面箍筋的实体

                BlockTableRecord lt = trans.GetObject(dwg.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                if (lt == null) 
                    throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "块表记录未找到");

                foreach (ObjectId item in selection.Value.GetObjectIds())
                {
                    Entity e = trans.GetObject(item, OpenMode.ForWrite) as Entity;
                    if (e == null) continue;

                    // 将直线转化为多段线
                    Line l = e as Line;
                    if (l != null)
                    {
                        Polyline pl = new Polyline();
                        pl.Layer = l.Layer;
                        pl.LineWeight = l.LineWeight;
                        pl.Linetype = l.Linetype;
                        pl.Color = l.Color;

                        pl.AddVertexAt(0, new Point2d(l.StartPoint.ToArray()), 0, 0, 0);
                        pl.AddVertexAt(1, new Point2d(l.EndPoint.ToArray()), 0, 0, 0);
                        pl.ConstantWidth = 25;

                        lt.AppendEntity(pl);
                        trans.AddNewlyCreatedDBObject(pl, true);
                        l.Erase();
                    } 
                    else
                    {
                        // 多段线修改全局宽度
                        Polyline pline = e as Polyline;
                        if (pline != null)
                        {
                            pline.ConstantWidth = 25;
                        }
                    }

                }   // foreach (ObjectId item in selection.Value.GetObjectIds())

                trans.Commit();

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

            // 从DWG图纸中获取梁及钢筋数据
            trans = dwg.TransactionManager.StartTransaction();
            try
            {
                // 确定单根梁范围，保存在beamRange中
                PromptSelectionResult selection = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.AXIS_NAME)
                        }
                    )
                );

                beamRange = new List<double[]>();
                bool isExisted = false;
                foreach (ObjectId item in selection.Value.GetObjectIds())
                {
                    Entity e = trans.GetObject(item, OpenMode.ForRead) as Entity;
                    if (e == null) continue;

                    Line l = e as Line;     // 轴线均为直线，不是多段线
                    if (l != null)
                    {
                        isExisted = false;
                        foreach (double[] t in beamRange)
                        {
                            if (Math.Abs(l.StartPoint.Y - t[0]) < Utils.DOUBLE_EPS ||
                                Math.Abs(l.EndPoint.Y - t[0]) < Utils.DOUBLE_EPS)   // 轴线一般水平方向对齐，以此进行分组
                            {
                                isExisted = true;
                            }
                        }
                        if (!isExisted)
                        {
                            beamRange.Add(
                                new double[] {
                                    l.StartPoint.Y, l.EndPoint.Y
                                }
                            );
                        }
                    }

                }   // foreach (ObjectId item in selection.Value.GetObjectIds())

                beamRange.Sort(
                    (double[] t1, double[] t2) => t1[0].CompareTo(t2[0])        // 排序，结果为从下至上（由小到大）
                );

                // 确定单根梁内部范围，保存在beamLineRange中
                // 确定单根梁支座线，保存在beamLineVertical中
                beamLineRange = new List<double[]>();
                beamLineVertical = new List<List<BoundingPolyline>>();
                for (int i = 0; i < beamRange.Count; i++)
                {
                    beamLineRange.Add(new double[] { double.PositiveInfinity, double.NegativeInfinity });
                    beamLineVertical.Add(new List<BoundingPolyline>());
                }

                PromptSelectionResult beamSelction = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_NAME)
                        }
                    )
                );

                foreach (ObjectId beamLine in beamSelction.Value.GetObjectIds())
                {
                    Line bl = trans.GetObject(beamLine, OpenMode.ForRead) as Line;
                    if (bl != null)
                    {
                        int blIndex = beamRange.FindIndex(
                            item => item[0] < bl.StartPoint.Y && bl.StartPoint.Y < item[1]
                        );  // 查找梁号，下同
                        if (blIndex == -1) continue;    // 未找到
                        Vector3d deltaVector = bl.EndPoint - bl.StartPoint;
                        if (Math.Abs(deltaVector.Y) <= Utils.DOUBLE_EPS)    // 为水平线
                        {
                            beamLineRange[blIndex][0] = Math.Min(beamLineRange[blIndex][0], bl.StartPoint.Y);
                            beamLineRange[blIndex][1] = Math.Max(beamLineRange[blIndex][1], bl.EndPoint.Y);
                        }
                    }
                }

                foreach (ObjectId beamLine in beamSelction.Value.GetObjectIds())
                {
                    Line bl = trans.GetObject(beamLine, OpenMode.ForRead) as Line;
                    if ( bl!= null)
                    {
                        int blIndex = beamRange.FindIndex(
                            item => item[0] < bl.StartPoint.Y && bl.StartPoint.Y < item[1]
                        );
                        if (blIndex == -1) continue;
                        Vector3d deltaVector = bl.EndPoint - bl.StartPoint;
                        // 为竖直线，且全部包含在梁水平线围合空间中
                        if ((Math.Abs(deltaVector.X) <= Utils.DOUBLE_EPS) 
                            && (beamLineRange[blIndex][0] - Utils.DOUBLE_EPS <= bl.StartPoint.Y)
                            && (bl.EndPoint.Y <= beamLineRange[blIndex][1] + Utils.DOUBLE_EPS)) 
                        {
                            beamLineVertical[blIndex].Add(new BoundingPolyline(bl));
                        }
                    }
                }

                // 创建钢筋表，将所有钢筋保存在steelBarTable中
                // 钢筋分类，保存在steelBarTableBeam及steelBarTableBeamReverse中
                steelBarTable = new List<BoundingPolyline>();
                steelBarTableBeam = new Dictionary<int, List<int>>();
                steelBarTableBeamReverse = new Dictionary<int, int>();

                for (int i = 0; i < beamRange.Count; i++)
                {
                    steelBarTableBeam.Add(
                        i, new List<int>()
                    );
                }

                PromptSelectionResult barSelection = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_STEEL_NAME)
                        }
                    )
                );

                foreach (ObjectId bar in barSelection.Value.GetObjectIds())
                {
                    Polyline pl = trans.GetObject(bar, OpenMode.ForRead) as Polyline;   // 钢筋均为多段线，没有直线
                    if (pl != null)
                    {
                        BoundingPolyline bp = new BoundingPolyline(pl);
                        int bpIndex = beamRange.FindIndex(
                            item => item[0] < bp.GetStartPoint().Y && bp.GetStartPoint().Y < item[1]
                        );      // 查找钢筋所在梁号
                        if (bpIndex == -1) continue;
                        Vector3d deltaVector = bp.GetEndPoint() - bp.GetStartPoint();
                        if (Math.Abs(deltaVector.Y) <= Utils.DOUBLE_EPS)    // 只考虑水平筋（锚固长度）
                        {
                            // 添加进钢筋表、梁-钢筋表、及其反查表
                            steelBarTableBeam[bpIndex].Add(steelBarTable.Count);
                            steelBarTableBeamReverse.Add(steelBarTable.Count, bpIndex);
                            steelBarTable.Add(bp);

                            // 判断钢筋是否在梁内
                            bp.isOut = true;
                            if ((beamLineRange[bpIndex][0] - Utils.DOUBLE_EPS <= bp.GetStartPoint().Y)
                                && (bp.GetEndPoint().Y <= beamLineRange[bpIndex][1] + Utils.DOUBLE_EPS))
                            {
                                bp.isOut = false;       // 如果钢筋位于梁水平线围合空间中，认为钢筋在梁内
                            }
                        }
                    }

                }   // foreach (ObjectId bar in barSelection.Value.GetObjectIds())

                // 计算水平筋锚固长度，并保存在
                // steelBarTableStart, steelBarTableStartReverse
                // steelBarTableEnd, steelBarTableEndReverse中
                steelBarTableStart = new Dictionary<double, List<int>>();
                steelBarTableEnd = new Dictionary<double, List<int>>();
                steelBarTableStartReverse = new Dictionary<int, double>();
                steelBarTableEndReverse = new Dictionary<int, double>();

                for (int i = 0; i < beamRange.Count; i++)
                {
                    foreach (int lineIndex in steelBarTableBeam[i])
                    {
                        // 计算锚固长度
                        double startMLength = double.PositiveInfinity;
                        double endMLength = double.PositiveInfinity;
                        BoundingPolyline l1 = steelBarTable[lineIndex];
                        foreach (BoundingPolyline l2 in beamLineVertical[i])
                        {
                            Point3dCollection pTemp = new Point3dCollection();
                            l1.GetLine().IntersectWith(
                                l2.GetLine(), Intersect.ExtendArgument, 
                                pTemp, IntPtr.Zero, IntPtr.Zero
                            );
                            if (pTemp.Count == 1)   // 有且仅有一个交点
                            {
                                startMLength = Math.Min(
                                    pTemp[0].DistanceTo(l1.GetStartPoint()),
                                    startMLength
                                );
                                endMLength = Math.Min(
                                    pTemp[0].DistanceTo(l1.GetEndPoint()),
                                    endMLength
                                );
                            }
                        }

                        // 添加进起始锚固长度表
                        if (steelBarTableStart.ContainsKey(Math.Round(startMLength, Utils.FIXED_NUM)))
                            steelBarTableStart[Math.Round(startMLength, Utils.FIXED_NUM)].Add(lineIndex);
                        else
                            steelBarTableStart.Add(
                                Math.Round(startMLength, Utils.FIXED_NUM),
                                new List<int> { lineIndex }
                            );

                        // 添加进末尾锚固长度表
                        if (steelBarTableEnd.ContainsKey(Math.Round(endMLength, Utils.FIXED_NUM)))
                            steelBarTableEnd[Math.Round(endMLength, Utils.FIXED_NUM)].Add(lineIndex);
                        else
                            steelBarTableEnd.Add(
                                Math.Round(endMLength, Utils.FIXED_NUM),
                                new List<int> { lineIndex }
                            );

                        // 添加进对应反查表
                        steelBarTableStartReverse.Add(
                            lineIndex,
                            Math.Round(startMLength, Utils.FIXED_NUM)
                        );
                        steelBarTableEndReverse.Add(
                            lineIndex,
                            Math.Round(endMLength, Utils.FIXED_NUM)
                        );
                    }   // foreach (int lineIndex in steelBarTableBeam[i])

                }   // for (int i = 0; i < beamRange.Count; i++)

                // 获取钢筋标注信息（根数、直径）
                PromptSelectionResult annoSelection = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_ANNO_NAME)
                        }
                    )
                );

                // 按梁号分类标注，提升计算效率
                List<List<DBText>> annoInBeam = new List<List<DBText>>();
                for (int i = 0; i < beamRange.Count; i++)
                {
                    annoInBeam.Add(new List<DBText>());
                }

                foreach (ObjectId anno in annoSelection.Value.GetObjectIds())
                {
                    DBText t = trans.GetObject(anno, OpenMode.ForRead) as DBText;
                    if (t != null)
                    {
                        int tIndex = beamRange.FindIndex(
                            item => item[0] < t.Position.Y && t.Position.Y < item[1]
                        );
                        if (tIndex == -1) continue;
                        annoInBeam[tIndex].Add(t);
                    }
                }

                // 分钢筋计算最近标注
                for (int i = 0; i < beamRange.Count; i++)
                {
                    foreach (int outLine in steelBarTableBeam[i])
                    {
                        if(steelBarTable[outLine].isOut)    // 外部钢筋才有相应标注
                        {
                            Point3d start = steelBarTable[outLine].GetStartPoint();
                            Point3d end = steelBarTable[outLine].GetEndPoint();
                            double distance = double.PositiveInfinity;
                            DBText ans = null;
                            foreach (DBText text in annoInBeam[i])
                            {
                                Point3d textPoint = text.Position;
                                if (start.X <= textPoint.X && textPoint.X <= end.X)     // 标注水平坐标在钢筋范围内
                                {
                                    double dis = Math.Abs(textPoint.Y - start.Y);
                                    if (dis < distance)
                                    {
                                        distance = dis;
                                        ans = text;                                     // 标注竖直距离最短
                                    }
                                }
                            }
                            // 如果有钢筋长度极短(e.g. <200mm)，可能会找不到标注
                            // 这种情况也无需考虑钢筋根数问题，该情况只可能出现在悬臂梁上部钢筋中
                            // 因此本程序选择跳过这种情况
                            if (ans == null) continue;  

                            // 利用正则表达式匹配钢筋数目和钢筋直径
                            Match match = Regex.Match(ans.TextString, Utils.ANNO_PATTERN);
                            if (match != null)
                            {
                                steelBarTable[outLine].num = int.Parse(match.Groups[1].Value);
                                steelBarTable[outLine].diameter = int.Parse(match.Groups[2].Value);
                            }
                        }   // if(steelBarTable[outLine].isOut) 

                    }   // foreach (int outLine in steelBarTableBeam[i])

                }   // for (int i = 0; i < beamRange.Count; i++)

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

            // 计算截面处钢筋标注
            trans = dwg.TransactionManager.StartTransaction();
            try
            {
                PromptSelectionResult steelBarAll = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_STEEL_NAME)
                        }
                    )
                );          // 所有纵筋

                PromptSelectionResult annoAll = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_ANNO_NAME)
                        }
                    )
                );          // 所有标注

                PromptSelectionResult annoName = ed.SelectAll(
                    new SelectionFilter(
                        new TypedValue[] {
                            new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_N_NAME)
                        }
                    )
                );          // 所有截面名称（获取截面分隔用）

                // 计算每个截面的X范围和Y范围
                profileRange = new List<double[]>();
                profileSplits = new List<List<double[]>>();
                List<List<DBText>> profileAnno = new List<List<DBText>>();      // 按梁重新组合标注

                for (int i = 0; i < beamRange.Count; i++)
                {
                    if (i != 0)
                    {
                        // 其余截面均为两组轴线围合范围
                        profileRange.Add(new double[] { beamRange[i - 1][1], beamRange[i][0] });
                    } 
                    else
                    {
                        // 第一个截面Y范围可到负无穷
                        profileRange.Add(new double[] { double.NegativeInfinity, beamRange[0][0] });
                    }
                    profileAnno.Add(new List<DBText>());
                    profileSplits.Add(new List<double[]>());
                }

                // 匹配所有截面标注，类似1-1, 2-2, 3-3 …
                foreach (ObjectId anno in annoName.Value.GetObjectIds())
                {
                    DBText text = trans.GetObject(anno, OpenMode.ForRead) as DBText;
                    if (text != null)
                    {
                        int textIndex = profileRange.FindIndex(
                            (double[] t) => t[0] <= text.Position.Y && text.Position.Y <= t[1]
                        );
                        if (textIndex == -1) continue;

                        Match match = Regex.Match(text.TextString, Utils.PROFILE_ANNO);
                        if (match != null && match.Groups.Count == 2)
                            profileAnno[textIndex].Add(text);
                    }
                }
                
                // 通过匹配的截面标注X坐标划分截面范围，取中点为分划
                // 这可能会导致划分截面偏左，但是对于读取钢筋可以接受
                for (int i = 0; i < profileRange.Count; i++)
                {
                    List<double> XTemp = new List<double>();
                    foreach (DBText text in profileAnno[i])
                    {
                        XTemp.Add(text.AlignmentPoint.X);
                    }
                    XTemp.Sort();   // 排序所有截面标注X坐标
                    if (XTemp.Count == 1)
                    {
                        // 如果只有一个截面信息，那么该截面X范围可为无穷
                        profileSplits[i].Add(new double[] { double.NegativeInfinity, double.PositiveInfinity });    
                        continue;
                    }
                    for (int j = 0; j < XTemp.Count - 1; j++)
                    {
                        if (j != 0)
                        {
                            // 一般情况为两个中点所围合的X范围
                            profileSplits[i].Add(
                                new double[] {
                                    (XTemp[j - 1] + XTemp[j]) / 2, 
                                    (XTemp[j] + XTemp[j + 1]) / 2
                                }
                            );
                        }
                        else
                        {
                            // 第一个截面的范围从负无穷开始
                            profileSplits[i].Add(
                                new double[] { 
                                    double.NegativeInfinity, 
                                    (XTemp[0] + XTemp[1]) / 2 
                                }
                            );
                        }

                    }   // for (int j = 0; j < XTemp.Count - 1; j++)

                    // 最后一个截面的范围可到正无穷
                    profileSplits[i].Add(
                        new double[] {
                            (XTemp[XTemp.Count - 2] + XTemp[XTemp.Count - 1]) / 2,
                            double.PositiveInfinity
                        }
                    );

                }   // for (int i = 0; i < profileRange.Count; i++)

                // 按梁-截面分类的标注信息
                List<List<List<DBText>>> beamSectionAnnos = new List<List<List<DBText>>>();
                // 按梁-截面分类的标注线
                List<List<List<BoundingPolyline>>> beamSectionAnnoLines = new List<List<List<BoundingPolyline>>>();
                // 按梁-截面分类的标注引线
                List<List<List<BoundingPolyline>>> beamSectionConnectLines = new List<List<List<BoundingPolyline>>>();

                steelBarBeamProfileRows = new Dictionary<int, Dictionary<int, Dictionary<double, List<BoundingPolyline>>>>();
                steelBarBeamProfileAnno = new Dictionary<int, Dictionary<int, Dictionary<string, string>>>();

                // 初始化这些三维/四维数组，预分配梁-截面两个索引的内存空间
                for (int i = 0; i < /* 梁个数= */profileRange.Count; i++)
                {
                    beamSectionAnnos.Add(new List<List<DBText>>());
                    beamSectionAnnoLines.Add(new List<List<BoundingPolyline>>());
                    beamSectionConnectLines.Add(new List<List<BoundingPolyline>>());
                    steelBarBeamProfileRows.Add(
                        i, new Dictionary<int, Dictionary<double, List<BoundingPolyline>>>()
                    );
                    steelBarBeamProfileAnno.Add(
                        i, new Dictionary<int, Dictionary<string, string>>()
                    );
                    for (int j = 0; j < /* 该梁对应截面个数= */profileSplits[i].Count; j++)
                    {
                        beamSectionAnnos[i].Add(new List<DBText>());
                        beamSectionAnnoLines[i].Add(new List<BoundingPolyline>());
                        beamSectionConnectLines[i].Add(new List<BoundingPolyline>());
                        steelBarBeamProfileRows[i].Add(j + 1, new Dictionary<double, List<BoundingPolyline>>());
                        steelBarBeamProfileAnno[i].Add(j + 1, new Dictionary<string, string>());
                    }

                }   // for (int i = 0; i < /* 梁个数= */profileRange.Count; i++)

                // 遍历所有标注、标注线、引线，按梁-截面分类保存
                foreach (ObjectId item in annoAll.Value.GetObjectIds())
                {
                    Entity e = trans.GetObject(item, OpenMode.ForRead) as Entity;
                    if (e == null) continue;

                    DBText text = e as DBText;
                    if (text != null)
                    {
                        // 计算梁号
                        int beamIndex = profileRange.FindIndex(
                            (double[] t) => t[0] <= text.Position.Y && text.Position.Y <= t[1]
                        );
                        if (beamIndex == -1) continue;
                        // 计算截面号
                        int profileIndex = profileSplits[beamIndex].FindIndex(
                            (double[] t) => t[0] <= text.Position.X && text.Position.X <= t[1]
                        );
                        if (profileIndex == -1) continue;
                        // 保存
                        beamSectionAnnos[beamIndex][profileIndex].Add(text);
                    }

                    Line l = e as Line;
                    if (l != null)
                    {
                        // 计算梁号
                        int beamIndex = profileRange.FindIndex(
                            (double[] t) => t[0] <= l.StartPoint.Y && l.StartPoint.Y <= t[1]
                        );
                        if (beamIndex == -1) continue;
                        // 计算截面号
                        int profileIndex = profileSplits[beamIndex].FindIndex(
                            (double[] t) => t[0] <= l.StartPoint.X && l.StartPoint.X <= t[1]
                        );
                        if (profileIndex == -1) continue;
                        Vector3d deltaVector = l.EndPoint - l.StartPoint;
                        if (Math.Abs(deltaVector.Y) <= Utils.DOUBLE_EPS)
                        {
                            // 认为标注线为水平线
                            beamSectionAnnoLines[beamIndex][profileIndex].Add(new BoundingPolyline(l));
                        }
                        else if (Math.Abs(deltaVector.X) <= Utils.DOUBLE_EPS)
                        {
                            // 认为标注引线为竖直线
                            beamSectionConnectLines[beamIndex][profileIndex].Add(new BoundingPolyline(l));
                        }

                    }   // if (l != null)

                }   //foreach (ObjectId item in annoAll.Value.GetObjectIds())

                // 按梁-截面分批遍历所有标注、引线、标注线，
                // 计算钢筋直径信息，并将钢筋分行保存在
                // steelBarBeamProfileRows中，
                // 同时计算标注信息，保存在
                // steelBarBeamProfileAnno中
                for (int i = 0; i < profileRange.Count; i++)
                {
                    for (int j = 0; j < profileSplits[i].Count; j++)
                    {
                        // 计算标注线所代表的标注值（最近标注）
                        foreach (BoundingPolyline annoLine in beamSectionAnnoLines[i][j])
                        {
                            Point3d start = annoLine.GetStartPoint();
                            Point3d end = annoLine.GetEndPoint();
                            double distance = double.PositiveInfinity;
                            DBText ans = null;
                            foreach (DBText text in beamSectionAnnos[i][j])
                            {
                                Point3d textPoint = text.Position;
                                if (annoLine.GetMinPointX() <= textPoint.X
                                    && textPoint.X <= annoLine.GetMaxPointX())     // 标注水平坐标在钢筋范围内
                                {
                                    double dis = textPoint.Y - start.Y;
                                    if (0 < dis && dis < distance)
                                    {
                                        distance = dis;
                                        ans = text;                                // 标注竖直距离最短
                                    }
                                }
                            }
                            // 有可能存在引线找不到标注，这主要发生在梁号的标注上
                            // 可能被划分成了两个不同的截面位置，导致找不到标注
                            // 但由于本程序不需要读取梁的截面信息，因此此处可以直接跳过
                            if (ans == null) continue;

                            // 利用正则表达式匹配普通钢筋钢筋数目和钢筋直径
                            Match match = Regex.Match(ans.TextString, Utils.ANNO_PATTERN);
                            if (match != null && match.Groups.Count == 3)
                            {
                                annoLine.num = int.Parse(match.Groups[1].Value);
                                annoLine.diameter = int.Parse(match.Groups[2].Value);
                            }
                            else
                            {
                                // 匹配架立筋或不伸入支座钢筋
                                Match matchJ = Regex.Match(ans.TextString, Utils.ANNO_PATTERN_J);
                                if (matchJ != null && matchJ.Groups.Count == 3)
                                {
                                    annoLine.isJ = true;
                                    annoLine.num = int.Parse(matchJ.Groups[1].Value);
                                    annoLine.diameter = int.Parse(matchJ.Groups[2].Value);
                                }
                            }
                        }   // foreach (BoundingPolyline annoLine in beamSectionAnnoLines[i][j])

                        // 计算标注引线所代表的标注值（末尾连接）
                        foreach (BoundingPolyline connectLine in beamSectionConnectLines[i][j])
                        {
                            Point3d start = connectLine.GetStartPoint();
                            Point3d end = connectLine.GetEndPoint();
                            foreach (BoundingPolyline annoLine in beamSectionAnnoLines[i][j])
                            {
                                // 标注引线末尾应在标注线上，如果不在，则为腰筋的标注，此处不进行考虑
                                if (end.Y == annoLine.GetStartPoint().Y)
                                {
                                    if (annoLine.GetMinPointX() - Utils.DOUBLE_EPS <= end.X &&
                                        end.X <= annoLine.GetMaxPointX())
                                    {
                                        connectLine.isJ = annoLine.isJ;
                                        connectLine.num = 1;
                                        connectLine.diameter = annoLine.diameter;
                                    }
                                }
                            }
                        }

                        // 计算钢筋线标注情况（引线起点在钢筋线中点处）
                        List<BoundingPolyline> steelLines = new List<BoundingPolyline>();
                        foreach (ObjectId steelBar in steelBarAll.Value.GetObjectIds())
                        {
                            Polyline pl = trans.GetObject(steelBar, OpenMode.ForRead) as Polyline;
                            if (pl != null)
                            {
                                // 计算钢筋线中心点坐标
                                Point3d midPoint = pl.GetLineSegmentAt(0).StartPoint
                                    + (pl.GetLineSegmentAt(0).EndPoint - pl.GetLineSegmentAt(0).StartPoint) / 2;

                                foreach (BoundingPolyline connectLine in beamSectionConnectLines[i][j])
                                {
                                    // 引线起点为钢筋线中点，并且引线不为腰筋引线
                                    if (midPoint == connectLine.GetStartPoint()
                                        && connectLine.diameter != -1)
                                    {
                                        BoundingPolyline bp = new BoundingPolyline(pl);
                                        bp.isJ = connectLine.isJ;
                                        bp.num = connectLine.num;
                                        bp.diameter = connectLine.diameter;
                                        steelLines.Add(bp);
                                    }
                                }

                            }   // if (pl != null)

                        }   // foreach (ObjectId steelBar in steelBarAll.Value.GetObjectIds())

                        // 对每个已标注的钢筋线，按行进行保存
                        foreach (BoundingPolyline bp in steelLines)
                        {
                            if (steelBarBeamProfileRows[i][j + 1].ContainsKey(Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM)))
                            {
                                steelBarBeamProfileRows[i][j + 1][Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM)].Add(bp);
                            }
                            else
                            {
                                steelBarBeamProfileRows[i][j + 1].Add(
                                    Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM),
                                    new List<BoundingPolyline> { bp }
                                );
                            }
                        }
                        
                        // 按从上至下进行排序（倒序），并通过中间Y值，划分上下配筋
                        List<double> steelBarYs = steelBarBeamProfileRows[i][j + 1].Keys.ToList();
                        steelBarYs.Sort();
                        steelBarYs.Reverse();

                        double midY = (steelBarYs.First() + steelBarYs.Last()) / 2;

                        List<Dictionary<int, int>> UpSteel = new List<Dictionary<int, int>>();
                        List<Dictionary<int, int>> DownSteel = new List<Dictionary<int, int>>();
#if DEBUG
                        ed.WriteMessage("\n梁号 {0}, 截面 {1}: ", i, j + 1);
#endif
                        foreach (double Y in steelBarYs)
                        {
                            // 下配筋Y较小
                            if (Y < midY)
                            {
                                DownSteel.Add(new Dictionary<int, int>());
                                foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])
                                {
                                    // 如果是架立筋，则将直径设为负值，下同
                                    if (b.isJ)
                                    {
                                        if (DownSteel.Last().ContainsKey(-b.diameter))
                                        {
                                            DownSteel.Last()[-b.diameter]++;
                                        }
                                        else
                                        {
                                            DownSteel.Last().Add(-b.diameter, 1);
                                        }
                                    }
                                    else
                                    {
                                        if (DownSteel.Last().ContainsKey(b.diameter))
                                        {
                                            DownSteel.Last()[b.diameter]++;
                                        }
                                        else
                                        {
                                            DownSteel.Last().Add(b.diameter, 1);
                                        }
                                    }

                                }   // foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])

                            }   // if (Y < midY)

                            // 上配筋Y较大
                            else
                            {
                                UpSteel.Add(new Dictionary<int, int>());
                                foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])
                                {
                                    if (b.isJ)
                                    {
                                        if (UpSteel.Last().ContainsKey(-b.diameter))
                                        {
                                            UpSteel.Last()[-b.diameter]++;
                                        }
                                        else
                                        {
                                            UpSteel.Last().Add(-b.diameter, 1);
                                        }
                                    }
                                    else
                                    {
                                        if (UpSteel.Last().ContainsKey(b.diameter))
                                        {
                                            UpSteel.Last()[b.diameter]++;
                                        }
                                        else
                                        {
                                            UpSteel.Last().Add(b.diameter, 1);
                                        }
                                    }

                                }   // foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])

                            }   // else

                        }   // foreach (double Y in steelBarYs)
#if DEBUG
                        ed.WriteMessage("上配筋: ");
#endif
                        // 生成上配筋字符串，保存在steelBarBeamProfileAnno中
                        string upSteelString = null;
                        bool isUnique = true;                               // 是否只有一个直径
                        int keyPrevious = UpSteel.First().First().Key;      // 第一个钢筋直径

                        foreach (var u in UpSteel)
                        {
                            if (u.Count == 1 && u.First().Key == keyPrevious) continue;
                            isUnique = false;
                        }
                        
                        if (isUnique)
                        {
                            // 如果只有一个钢筋直径，则可以进行简写 例如: "8Ф25 6/2"
                            List<int> rowCount = new List<int>();
                            foreach (var u in UpSteel)
                            {
                                rowCount.Add(u.First().Value);
                            }
                            if (rowCount.Count == 1)
                            {
                                upSteelString = $"{rowCount.Sum()}%%132{UpSteel.First().First().Key}";
                            }
                            else
                            {
                                upSteelString = $"{rowCount.Sum()}%%132{UpSteel.First().First().Key} {string.Join("/", rowCount)}";
                            }
                        } 
                        else
                        {
                            // 如果有多个钢筋直径，则不能简写，注意"+"连接同一行钢筋
                            // "/"分隔不同行钢筋，架立筋用括号括起
                            List<string> rowSteelText = new List<string>();
                            foreach (var u in UpSteel)
                            {
                                List<string> record = new List<string>();
                                foreach (var k in u)
                                {
                                    if (k.Key > 0)
                                    {
                                        record.Add($"{k.Value}%%132{k.Key}");
                                    }
                                    else
                                    {
                                        // 架立筋用括号括起
                                        record.Add($"({k.Value}%%132{-k.Key})");
                                    }
                                }
                                // "+"连接同一行钢筋
                                rowSteelText.Add(string.Join("+", record));
                            }
                            // "/"分隔不同行钢筋
                            upSteelString = string.Join("/", rowSteelText);

                        }   // else

                        steelBarBeamProfileAnno[i][j + 1].Add("UP", upSteelString);

#if DEBUG
                        ed.WriteMessage(upSteelString);
                        ed.WriteMessage(", 下配筋: ");
#endif
                        // 生成下配筋字符串，保存在steelBarBeamProfileAnno中
                        string downSteelString = null;
                        isUnique = true;                                            // 是否只有一个直径
                        keyPrevious = Math.Abs(DownSteel.First().First().Key);      // 第一个钢筋直径

                        foreach (var u in DownSteel)
                        {
                            foreach (int k in u.Keys)
                            {
                                if (Math.Abs(k) == keyPrevious) continue;
                                isUnique = false;
                            }
                        }

                        if (isUnique)
                        {
                            // 如果只有一个钢筋直径，则需要考虑不伸入支座钢筋情况
                            // 此处用两个列表保存
                            List<int> rowCount = new List<int>();               // 存储总钢筋数
                            List<int> rowJCount = new List<int>();              // 存储不伸入支座钢筋数量
                            foreach (var u in DownSteel)
                            {
                                // 如果该行仅有一种钢筋
                                if (u.Count == 1)                               
                                {
                                    // 本行全部不是不伸入支座钢筋
                                    if (u.First().Key > 0)                      
                                    {
                                        rowCount.Add(u.First().Value);
                                        rowJCount.Add(0);
                                    }
                                    // 本行全部为不伸入支座钢筋
                                    else
                                    {
                                        rowCount.Add(u.First().Value);
                                        rowJCount.Add(-u.First().Value);
                                    }
                                }
                                // 如果该行有两种钢筋，注意到此时仅有一种钢筋直径
                                // 说明u中仅有两个元素，一个是不伸入支座钢筋，一个贯通
                                else
                                {
                                    // 如果第一个钢筋贯通
                                    if (u.First().Key > 0)                      
                                    {
                                        rowCount.Add(u.First().Value + u.Last().Value);
                                        // 第二个为伸入支座钢筋
                                        rowJCount.Add(-u.Last().Value);         
                                    }
                                    // 如果第一个钢筋不伸入支座
                                    else
                                    {
                                        rowCount.Add(u.First().Value + u.Last().Value);
                                        rowJCount.Add(-u.First().Value);
                                    }
                                }

                            }   // foreach (var u in DownSteel)

                            // 如果仅一行
                            if (rowCount.Count == 1)                            
                            {
                                // 如果没有不伸入支座钢筋
                                if (rowJCount.First() == 0)
                                {
                                    downSteelString = $"{rowCount.Sum()}%%132{DownSteel.First().First().Key}";
                                }
                                // 如果存在不伸入支座钢筋
                                else
                                {
                                    downSteelString = $"{rowCount.Sum()}%%132{DownSteel.First().First().Key}({rowJCount.Sum()})";
                                }
                            }
                            // 如果不仅一行
                            else
                            {
                                List<string> appendixString = new List<string>();
                                for (int k = 0; k < rowCount.Count; k++)
                                {
                                    // 该行无不伸入支座钢筋
                                    if (rowJCount[k] == 0)
                                    {
                                        appendixString.Add($"{rowCount[k]}");
                                    }
                                    // 该行有不伸入支座钢筋，减少数量标在括号内
                                    else
                                    {
                                        appendixString.Add($"{rowCount[k]}({rowJCount[k]})");
                                    }
                                }
                                // 组合上述标注
                                downSteelString = $"{rowCount.Sum()}%%132{keyPrevious} {string.Join("/", appendixString)}";
                            } 

                        }   // if (isUnique)

                        // 如果不仅一种钢筋直径
                        else
                        {
                            List<string> rowSteelText = new List<string>();
                            foreach (var u in DownSteel)
                            {
                                // 一条记录，例如"8Ф25(-2)"
                                List<string> record = new List<string>();
                                foreach (var k in u)
                                {
                                    // 如果k是不伸入支座的钢筋
                                    if (k.Key < 0)
                                    {
                                        // 如果该行同时存在相同直径的伸入支座钢筋
                                        if (u.ContainsKey(-k.Key))
                                        {
                                            // 留到伸入支座钢筋处理
                                            continue;
                                        }
                                        // 如果该行不存在相同直径的伸入支座钢筋
                                        else
                                        {
                                            record.Add($"{k.Value}%%132{-k.Key}(-{k.Value})");
                                        }
                                    }
                                    // 如果k是伸入支座的钢筋
                                    else
                                    {
                                        // 如果该行同时存在相同直径的不伸入支座钢筋
                                        if (u.ContainsKey(-k.Key))
                                        {
                                            record.Add($"{k.Value + u[-k.Key]}%%132{k.Key}(-{u[-k.Key]})");
                                        }
                                        // 如果该行不存在相同直径的不伸入支座钢筋
                                        else
                                        {
                                            record.Add($"{k.Value}%%132{k.Key}");
                                        }
                                    }

                                }   // foreach (var k in u)

                                rowSteelText.Add(string.Join("+", record));

                            }   // foreach (var u in DownSteel)

                            downSteelString = string.Join("/", rowSteelText);

                        }   // else

                        steelBarBeamProfileAnno[i][j + 1].Add("DOWN", downSteelString);
#if DEBUG
                        ed.WriteMessage(downSteelString);
#endif
                    }   // for (int j = 0; j < profileSplits[i].Count; j++)

                }   // for (int i = 0; i < profileRange.Count; i++)

                // 输出基本信息
                ed.WriteMessage(
                    "\n图中检测出 " +
                    beamRange.Count.ToString() +
                    " 条梁"
                );
                int outBarCount, inBarCount;
                foreach (var k in steelBarTableBeam.Keys)
                {
                    outBarCount = 0; inBarCount = 0;
                    foreach (var m in steelBarTableBeam[k])
                    {
                        if (steelBarTable[m].isOut) outBarCount++;
                        else inBarCount++;
                    }
                    ed.WriteMessage(
                        "\n梁号 {0}, 截面数量 {1}, 总钢筋数量 {2}, 外部钢筋数量 {3}, 内部钢筋数量 {4}",
                        k, profileSplits[k].Count, steelBarTableBeam[k].Count,
                        outBarCount, inBarCount
                    );
                }
                ed.WriteMessage("\n总钢筋数 {0}", steelBarTableBeamReverse.Count);

            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

            // 询问是否进行自动标注，因为自动标注会产生一些冗余，因此进行询问
            PromptResult isAnnoResult = ed.GetKeywords(
                new PromptKeywordOptions("\n是否自动标注 [Y/N] : ", "Y N")
            );

            if (isAnnoResult.Status == PromptStatus.OK &&
                "Y".Equals(isAnnoResult.StringResult))
            {
                // 启动自动标注
                trans = dwg.TransactionManager.StartTransaction();
                try
                {
                    // 查找需要标注的位置及对应截面号
                    PromptSelectionResult profileSection = ed.SelectAll(
                        new SelectionFilter(
                            new TypedValue[] {
                                new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_ANNO_NAME)
                            }
                        )
                    );          // 选中所有截面标注

                    // 打开块表记录
                    BlockTableRecord btr = trans.GetObject(dwg.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr == null)
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "块表记录未找到");

                    // 使用"样式 1"进行标注
                    TextStyleTable txtStlTbl = trans.GetObject(dwg.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    ObjectId style = txtStlTbl[Utils.ANNO_STYLE_NAME];

                    // 遍历所有标注，查找截面标记
                    foreach (ObjectId item in profileSection.Value.GetObjectIds())
                    {
                        DBText text = trans.GetObject(item, OpenMode.ForRead) as DBText;
                        if (text == null) continue;

                        // 查找梁号
                        int beamIndex = beamRange.FindIndex(
                            (double[] t) => t[0] <= text.Position.Y && text.Position.Y <= t[1]
                        );
                        if (beamIndex == -1) continue;

                        // 过滤是否为截面标记，并获取截面号
                        Match match = Regex.Match(text.TextString, Utils.PROFILE_SECTION);
                        if (match == null) continue;
                        if (match.Groups.Count != 2) continue;

                        int profileIndex = int.Parse(match.Groups[1].Value);

                        // 此处取梁上部分标注，梁下部分标注舍去
                        if (text.Position.Y > beamLineRange[beamIndex][1])
                        {
                            // 新建梁上端标注
                            DBText newText = new DBText();
                            // 与截面标注同层
                            newText.Layer = text.Layer;
                            // 样式取为"样式 1"
                            newText.TextStyleId = style;
                            // 右对齐
                            newText.HorizontalMode = TextHorizontalMode.TextRight;
                            // 宽度因子0.7
                            newText.WidthFactor = 0.7;
                            // 字高150
                            newText.Height = 150;

                            newText.TextString = steelBarBeamProfileAnno[beamIndex][profileIndex]["UP"];
                            // 对齐点位置利用标注符号及梁上端线偏移一定距离
                            newText.AlignmentPoint = new Point3d(text.Position.X - 100, beamLineRange[beamIndex][1] + 45, 0);
                            // 添加标注
                            btr.AppendEntity(newText);
                            trans.AddNewlyCreatedDBObject(newText, true);

                            // 新建梁下端标注
                            DBText downText = new DBText();
                            downText.Layer = text.Layer;
                            downText.TextStyleId = style;
                            downText.VerticalMode = TextVerticalMode.TextTop;
                            downText.HorizontalMode = TextHorizontalMode.TextRight;
                            downText.WidthFactor = 0.7;
                            downText.Height = 150;
                            downText.TextString = steelBarBeamProfileAnno[beamIndex][profileIndex]["DOWN"];
                            // 对齐点位置利用标注符号及梁上端线偏移一定距离
                            downText.AlignmentPoint = new Point3d(text.Position.X - 100, beamLineRange[beamIndex][0] - 45, 0);

                            btr.AppendEntity(downText);
                            trans.AddNewlyCreatedDBObject(downText, true);

                        }   // if (text.Position.Y > beamLineRange[beamIndex][1])

                    }   // foreach (ObjectId item in profileSection.Value.GetObjectIds())

                    // 提交事务
                    trans.Commit();

                }   // try
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage("problem due to " + ex.Message);
                }
                finally
                {
                    trans.Dispose();
                }

            }   // if (isAnnoResult.Status == PromptStatus.OK && "Y".Equals(isAnnoResult.StringResult))

        }   // public void YJKParse()

        /// <summary>
        /// 计算单一截面上的钢筋情况
        /// </summary>
        [CommandMethod("K")]
        public void AnnoParse()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            Database dwg = ed.Document.Database;

            PromptSelectionResult profileAll = ed.GetSelection();

            if (profileAll.Status == PromptStatus.OK)
            {
                Transaction trans = dwg.TransactionManager.StartTransaction();

                try
                {
                    List<Entity> annoAll = new List<Entity>();
                    List<Entity> steelBarAll = new List<Entity>();

                    foreach (ObjectId item in profileAll.Value.GetObjectIds())
                    {
                        Entity e = trans.GetObject(item, OpenMode.ForRead) as Entity;
                        if (e == null) continue;

                        if (e.Layer.Equals(Utils.BEAM_SECTION_ANNO_NAME))
                        {
                            annoAll.Add(e);
                        }
                        else if (e.Layer.Equals(Utils.BEAM_SECTION_STEEL_NAME))
                        {
                            steelBarAll.Add(e);
                        }
                    }

                    List<BoundingPolyline> annoLines = new List<BoundingPolyline>();        // 标注线
                    List<BoundingPolyline> connectLines = new List<BoundingPolyline>();     // 引线
                    List<DBText> annos = new List<DBText>();                                // 钢筋标注

                    foreach (Entity anno in annoAll)
                    {
                        Line l = anno as Line;
                        if (l != null)
                        {
                            Vector3d deltaVector = l.EndPoint - l.StartPoint;
                            if (Math.Abs(deltaVector.Y) <= Utils.DOUBLE_EPS)
                            {
                                annoLines.Add(new BoundingPolyline(l));                     // 标注线是水平线
                            }
                            else if (Math.Abs(deltaVector.X) <= Utils.DOUBLE_EPS)
                            {
                                connectLines.Add(new BoundingPolyline(l));                  // 引线是垂直线
                            }
                        }
                        DBText t = anno as DBText;
                        if (t != null)
                        {
                            annos.Add(t);
                        }
                    }

                    foreach (BoundingPolyline annoLine in annoLines)
                    {
                        Point3d start = annoLine.GetStartPoint();
                        Point3d end = annoLine.GetEndPoint();
                        double distance = double.PositiveInfinity;
                        DBText ans = null;
                        foreach (DBText text in annos)
                        {
                            Point3d textPoint = text.Position;
                            if (annoLine.GetMinPointX() <= textPoint.X
                                && textPoint.X <= annoLine.GetMaxPointX())     // 标注水平坐标在钢筋范围内
                            {
                                double dis = textPoint.Y - start.Y;
                                if (0 < dis && dis < distance)
                                {
                                    distance = dis;
                                    ans = text;                                // 标注竖直距离最短
                                }
                            }
                        }

                        // 利用正则表达式匹配钢筋数目和钢筋直径
                        Match match = Regex.Match(ans.TextString, Utils.ANNO_PATTERN);
                        if (match != null && match.Groups.Count == 3)
                        {
                            annoLine.num = int.Parse(match.Groups[1].Value);
                            annoLine.diameter = int.Parse(match.Groups[2].Value);
                        }
                        else
                        {
                            Match matchJ = Regex.Match(ans.TextString, Utils.ANNO_PATTERN_J);
                            if (matchJ != null && matchJ.Groups.Count == 3)
                            {
                                annoLine.isJ = true;
                                annoLine.num = int.Parse(matchJ.Groups[1].Value);
                                annoLine.diameter = int.Parse(matchJ.Groups[2].Value);
                            }
                        }
                    }

                    // 计算引线代表的钢筋直径
                    foreach (BoundingPolyline connectLine in connectLines)
                    {
                        Point3d start = connectLine.GetStartPoint();
                        Point3d end = connectLine.GetEndPoint();
                        foreach (BoundingPolyline annoLine in annoLines)
                        {
                            if (end.Y == annoLine.GetStartPoint().Y)
                            {
                                if (annoLine.GetMinPointX() - Utils.DOUBLE_EPS <= end.X &&
                                    end.X <= annoLine.GetMaxPointX())
                                {
                                    connectLine.isJ = annoLine.isJ;
                                    connectLine.num = 1;
                                    connectLine.diameter = annoLine.diameter;
                                }
                            }
                        }
                    }

                    // 计算钢筋（圆形）的直径
                    List<BoundingPolyline> steelLines = new List<BoundingPolyline>();
                    foreach (Entity steelBar in steelBarAll)
                    {
                        Polyline pl = steelBar as Polyline;
                        if (pl != null)
                        {
                            Point3d midPoint = pl.GetLineSegmentAt(0).StartPoint
                                + (pl.GetLineSegmentAt(0).EndPoint - pl.GetLineSegmentAt(0).StartPoint) / 2; // 中心点坐标

                            foreach (BoundingPolyline connectLine in connectLines)
                            {
                                if (midPoint == connectLine.GetStartPoint()
                                    && connectLine.diameter != -1)
                                {
                                    BoundingPolyline bp = new BoundingPolyline(pl);
                                    bp.isJ = connectLine.isJ;
                                    bp.num = connectLine.num;
                                    bp.diameter = connectLine.diameter;
                                    steelLines.Add(bp);
                                }
                            }
                        }
                    }

                    // 将钢筋按行排列
                    Dictionary<double, List<BoundingPolyline>> steelBarRows
                        = new Dictionary<double, List<BoundingPolyline>>();
                    foreach (BoundingPolyline bp in steelLines)
                    {
                        if (steelBarRows.ContainsKey(Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM)))
                        {
                            steelBarRows[Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM)].Add(bp);
                        }
                        else
                        {
                            steelBarRows.Add(
                                Math.Round(bp.GetStartPoint().Y, Utils.FIXED_NUM),
                                new List<BoundingPolyline> { bp }
                            );
                        }
                    }

                    // 确定上部钢筋与下部钢筋
                    List<double> steelBarYs = steelBarRows.Keys.ToList();
                    steelBarYs.Sort();
                    steelBarYs.Reverse();
                    double midY = (steelBarYs.First() + steelBarYs.Last()) / 2;

                    List<Dictionary<int, int>> UpSteel = new List<Dictionary<int, int>>();
                    List<Dictionary<int, int>> DownSteel = new List<Dictionary<int, int>>();

                    foreach (double Y in steelBarYs)
                    {
                        // 下配筋Y较小
                        if (Y < midY)
                        {
                            DownSteel.Add(new Dictionary<int, int>());
                            foreach (var b in steelBarRows[Y])
                            {
                                // 如果是架立筋，则将直径设为负值，下同
                                if (b.isJ)
                                {
                                    if (DownSteel.Last().ContainsKey(-b.diameter))
                                    {
                                        DownSteel.Last()[-b.diameter]++;
                                    }
                                    else
                                    {
                                        DownSteel.Last().Add(-b.diameter, 1);
                                    }
                                }
                                else
                                {
                                    if (DownSteel.Last().ContainsKey(b.diameter))
                                    {
                                        DownSteel.Last()[b.diameter]++;
                                    }
                                    else
                                    {
                                        DownSteel.Last().Add(b.diameter, 1);
                                    }
                                }

                            }   // foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])

                        }   // if (Y < midY)

                        // 上配筋Y较大
                        else
                        {
                            UpSteel.Add(new Dictionary<int, int>());
                            foreach (var b in steelBarRows[Y])
                            {
                                if (b.isJ)
                                {
                                    if (UpSteel.Last().ContainsKey(-b.diameter))
                                    {
                                        UpSteel.Last()[-b.diameter]++;
                                    }
                                    else
                                    {
                                        UpSteel.Last().Add(-b.diameter, 1);
                                    }
                                }
                                else
                                {
                                    if (UpSteel.Last().ContainsKey(b.diameter))
                                    {
                                        UpSteel.Last()[b.diameter]++;
                                    }
                                    else
                                    {
                                        UpSteel.Last().Add(b.diameter, 1);
                                    }
                                }

                            }   // foreach (var b in steelBarBeamProfileRows[i][j + 1][Y])

                        }   // else

                    }   // foreach (double Y in steelBarYs)

                    ed.WriteMessage("上配筋: ");

                    // 生成上配筋字符串，保存在steelBarBeamProfileAnno中
                    string upSteelString = null;
                    bool isUnique = true;                               // 是否只有一个直径
                    int keyPrevious = UpSteel.First().First().Key;      // 第一个钢筋直径

                    foreach (var u in UpSteel)
                    {
                        if (u.Count == 1 && u.First().Key == keyPrevious) continue;
                        isUnique = false;
                    }

                    if (isUnique)
                    {
                        // 如果只有一个钢筋直径，则可以进行简写 例如: "8Ф25 6/2"
                        List<int> rowCount = new List<int>();
                        foreach (var u in UpSteel)
                        {
                            rowCount.Add(u.First().Value);
                        }
                        if (rowCount.Count == 1)
                        {
                            upSteelString = $"{rowCount.Sum()}Φ{UpSteel.First().First().Key}";
                        }
                        else
                        {
                            upSteelString = $"{rowCount.Sum()}Φ{UpSteel.First().First().Key} {string.Join("/", rowCount)}";
                        }
                    }
                    else
                    {
                        // 如果有多个钢筋直径，则不能简写，注意"+"连接同一行钢筋
                        // "/"分隔不同行钢筋，架立筋用括号括起
                        List<string> rowSteelText = new List<string>();
                        foreach (var u in UpSteel)
                        {
                            List<string> record = new List<string>();
                            foreach (var k in u)
                            {
                                if (k.Key > 0)
                                {
                                    record.Add($"{k.Value}Φ{k.Key}");
                                }
                                else
                                {
                                    // 架立筋用括号括起
                                    record.Add($"({k.Value}Φ{-k.Key})");
                                }
                            }
                            // "+"连接同一行钢筋
                            rowSteelText.Add(string.Join("+", record));
                        }
                        // "/"分隔不同行钢筋
                        upSteelString = string.Join("/", rowSteelText);

                    }   // else

                    ed.WriteMessage(upSteelString);
                    ed.WriteMessage(", 下配筋: ");

                    // 生成下配筋字符串，保存在steelBarBeamProfileAnno中
                    string downSteelString = null;
                    isUnique = true;                                            // 是否只有一个直径
                    keyPrevious = Math.Abs(DownSteel.First().First().Key);      // 第一个钢筋直径

                    foreach (var u in DownSteel)
                    {
                        foreach (int k in u.Keys)
                        {
                            if (Math.Abs(k) == keyPrevious) continue;
                            isUnique = false;
                        }
                    }

                    if (isUnique)
                    {
                        // 如果只有一个钢筋直径，则需要考虑不伸入支座钢筋情况
                        // 此处用两个列表保存
                        List<int> rowCount = new List<int>();               // 存储总钢筋数
                        List<int> rowJCount = new List<int>();              // 存储不伸入支座钢筋数量
                        foreach (var u in DownSteel)
                        {
                            // 如果该行仅有一种钢筋
                            if (u.Count == 1)
                            {
                                // 本行全部不是不伸入支座钢筋
                                if (u.First().Key > 0)
                                {
                                    rowCount.Add(u.First().Value);
                                    rowJCount.Add(0);
                                }
                                // 本行全部为不伸入支座钢筋
                                else
                                {
                                    rowCount.Add(u.First().Value);
                                    rowJCount.Add(-u.First().Value);
                                }
                            }
                            // 如果该行有两种钢筋，注意到此时仅有一种钢筋直径
                            // 说明u中仅有两个元素，一个是不伸入支座钢筋，一个贯通
                            else
                            {
                                // 如果第一个钢筋贯通
                                if (u.First().Key > 0)
                                {
                                    rowCount.Add(u.First().Value + u.Last().Value);
                                    // 第二个为伸入支座钢筋
                                    rowJCount.Add(-u.Last().Value);
                                }
                                // 如果第一个钢筋不伸入支座
                                else
                                {
                                    rowCount.Add(u.First().Value + u.Last().Value);
                                    rowJCount.Add(-u.First().Value);
                                }
                            }

                        }   // foreach (var u in DownSteel)

                        // 如果仅一行
                        if (rowCount.Count == 1)
                        {
                            // 如果没有不伸入支座钢筋
                            if (rowJCount.First() == 0)
                            {
                                downSteelString = $"{rowCount.Sum()}Φ{DownSteel.First().First().Key}";
                            }
                            // 如果存在不伸入支座钢筋
                            else
                            {
                                downSteelString = $"{rowCount.Sum()}Φ{DownSteel.First().First().Key}({rowJCount.Sum()})";
                            }
                        }
                        // 如果不仅一行
                        else
                        {
                            List<string> appendixString = new List<string>();
                            for (int k = 0; k < rowCount.Count; k++)
                            {
                                // 该行无不伸入支座钢筋
                                if (rowJCount[k] == 0)
                                {
                                    appendixString.Add($"{rowCount[k]}");
                                }
                                // 该行有不伸入支座钢筋，减少数量标在括号内
                                else
                                {
                                    appendixString.Add($"{rowCount[k]}({rowJCount[k]})");
                                }
                            }
                            // 组合上述标注
                            downSteelString = $"{rowCount.Sum()}Φ{keyPrevious} {string.Join("/", appendixString)}";
                        }

                    }   // if (isUnique)

                    // 如果不仅一种钢筋直径
                    else
                    {
                        List<string> rowSteelText = new List<string>();
                        foreach (var u in DownSteel)
                        {
                            // 一条记录，例如"8Ф25(-2)"
                            List<string> record = new List<string>();
                            foreach (var k in u)
                            {
                                // 如果k是不伸入支座的钢筋
                                if (k.Key < 0)
                                {
                                    // 如果该行同时存在相同直径的伸入支座钢筋
                                    if (u.ContainsKey(-k.Key))
                                    {
                                        // 留到伸入支座钢筋处理
                                        continue;
                                    }
                                    // 如果该行不存在相同直径的伸入支座钢筋
                                    else
                                    {
                                        record.Add($"{k.Value}Φ{-k.Key}(-{k.Value})");
                                    }
                                }
                                // 如果k是伸入支座的钢筋
                                else
                                {
                                    // 如果该行同时存在相同直径的不伸入支座钢筋
                                    if (u.ContainsKey(-k.Key))
                                    {
                                        record.Add($"{k.Value + u[-k.Key]}Φ{k.Key}(-{u[-k.Key]})");
                                    }
                                    // 如果该行不存在相同直径的不伸入支座钢筋
                                    else
                                    {
                                        record.Add($"{k.Value}Φ{k.Key}");
                                    }
                                }

                            }   // foreach (var k in u)

                            rowSteelText.Add(string.Join("+", record));

                        }   // foreach (var u in DownSteel)

                        downSteelString = string.Join("/", rowSteelText);

                    }   // else

                    ed.WriteMessage(downSteelString);

                }   // try

                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage("problem due to " + ex.Message);
                }
                finally
                {
                    trans.Dispose();
                }

            }   // if (profileAll.Status == PromptStatus.OK)

        }   // public void AnnoParse()

        /// <summary>
        /// 手动标注
        /// </summary>
        [CommandMethod("S")]
        public void ManualAnno()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            if (steelBarBeamProfileAnno == null)
            {
                ed.WriteMessage("\n请先运行初始化命令 YJK \n");
                return;
            }

            Database dwg = ed.Document.Database;

            PromptPointResult annoPosition = ed.GetPoint("请选择标注位置");

            if (annoPosition.Status == PromptStatus.OK)
            {
                Transaction trans = dwg.TransactionManager.StartTransaction();
                try
                {
                    // 初始化profileSectionText
                    if (profileSectionText == null)
                    {
                        profileSectionText = new List<List<DBText>>();

                        for (int i = 0; i < beamRange.Count; i++)
                        {
                            profileSectionText.Add(new List<DBText>());
                        }

                        // 查找需要标注的位置及对应截面号
                        PromptSelectionResult profileSection = ed.SelectAll(
                            new SelectionFilter(
                                new TypedValue[] {
                                    new TypedValue((int)DxfCode.LayerName, Utils.BEAM_SECTION_ANNO_NAME)
                                }
                            )
                        );          // 选中所有截面标注

                        foreach (ObjectId item in profileSection.Value.GetObjectIds())
                        {
                            DBText text = trans.GetObject(item, OpenMode.ForRead) as DBText;
                            if (text == null) continue;

                            int textBeamIndex = beamRange.FindIndex(
                                (double[] t) => t[0] <= text.Position.Y && text.Position.Y <= t[1]
                            );
                            if (textBeamIndex == -1) continue;

                            Match match = Regex.Match(text.TextString, Utils.PROFILE_SECTION);
                            if (match == null) continue;
                            if (match.Groups.Count != 2) continue;

                            if (text.Position.Y > beamLineRange[textBeamIndex][1])
                                profileSectionText[textBeamIndex].Add(text);
                        }
                    }

                    int beamIndex = beamRange.FindIndex(
                        (double[] t) => t[0] <= annoPosition.Value.Y && annoPosition.Value.Y <= t[1]
                    );

                    // 打开块表记录
                    BlockTableRecord btr = trans.GetObject(dwg.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    if (btr == null)
                        throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "块表记录未找到");

                    // 使用"样式 1"进行标注
                    TextStyleTable txtStlTbl = trans.GetObject(dwg.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    ObjectId style = txtStlTbl[Utils.ANNO_STYLE_NAME];

                    DBText selectText = null;
                    double minDis = double.PositiveInfinity;

                    // 遍历所有标注，查找最近的截面标记
                    foreach (DBText text in profileSectionText[beamIndex])
                    {
                        double dis = Math.Abs(text.Position.X - annoPosition.Value.X);
                        if (dis < minDis)
                        {
                            minDis = dis;
                            selectText = text;
                        }
                    }

                    // 获取截面号
                    Match match1 = Regex.Match(selectText.TextString, Utils.PROFILE_SECTION);
                    int profileIndex = int.Parse(match1.Groups[1].Value);

                    // 如果选择点在梁上端
                    if (annoPosition.Value.Y > beamLineRange[beamIndex][1])
                    {
                        
                        DBText newText = new DBText();
                        // 与截面标注同层
                        newText.Layer = selectText.Layer;
                        // 样式取为"样式 1"
                        newText.TextStyleId = style;
                        // 右对齐
                        newText.HorizontalMode = TextHorizontalMode.TextRight;
                        // 宽度因子0.7
                        newText.WidthFactor = 0.7;
                        // 字高150
                        newText.Height = 150;

                        newText.TextString = steelBarBeamProfileAnno[beamIndex][profileIndex]["UP"];
                        // 对齐点位置即为选择点
                        newText.AlignmentPoint = annoPosition.Value;
                        // 添加标注
                        btr.AppendEntity(newText);
                        trans.AddNewlyCreatedDBObject(newText, true);

                    }   // if (annoPosition.Value.Y > beamLineRange[beamIndex][1])

                    // 选择点在梁下端
                    else
                    { 
                        DBText downText = new DBText();
                        downText.Layer = selectText.Layer;
                        downText.TextStyleId = style;
                        downText.VerticalMode = TextVerticalMode.TextTop;
                        downText.HorizontalMode = TextHorizontalMode.TextRight;
                        downText.WidthFactor = 0.7;
                        downText.Height = 150;
                        downText.TextString = steelBarBeamProfileAnno[beamIndex][profileIndex]["DOWN"];
                        // 对齐点位置为选择点
                        downText.AlignmentPoint = annoPosition.Value;

                        btr.AppendEntity(downText);
                        trans.AddNewlyCreatedDBObject(downText, true);

                    }   // else

                    // 提交事务
                    trans.Commit();

                }   // try
                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    ed.WriteMessage("problem due to " + ex.Message);
                }
                finally
                {
                    trans.Dispose();
                }
            }
           
        }   // public void ManualAnno()

        /// <summary>
        /// 腰筋自动标注
        /// </summary>
        [CommandMethod("O")]
        public void SteelBarDim()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            if (beamLineRange == null)
            {
                ed.WriteMessage("\n请先运行初始化命令 YJK \n");
                return;
            }

            Database dwg = ed.Document.Database;

            Transaction trans = dwg.TransactionManager.StartTransaction();
            try
            {
                for (int i = 0; i < beamRange.Count; i++)
                {
                    foreach (int steelIndex in steelBarTableBeam[i])
                    {
                        bool isExist = false;
                        foreach (int sameStartSteelIndex in steelBarTableStart[steelBarTableStartReverse[steelIndex]])
                        {
                            if (steelBarTable[sameStartSteelIndex].isOut
                                && steelBarTableBeamReverse[sameStartSteelIndex] == i)
                            {
                                isExist = true;
                                break;
                            }
                        }

                        foreach (int sameEndSteelIndex in steelBarTableEnd[steelBarTableEndReverse[steelIndex]])
                        {
                            if (steelBarTable[sameEndSteelIndex].isOut
                                && steelBarTableBeamReverse[sameEndSteelIndex] == i)
                            {
                                isExist = true;
                                break;
                            }
                        }
                        
                        if (isExist) continue;
                        // 获取所有的腰筋
                        BoundingPolyline l1 = steelBarTable[steelIndex];

                        double startMLength = double.PositiveInfinity;
                        double endMLength = double.PositiveInfinity;
                        BoundingPolyline lLeft = null;
                        BoundingPolyline lRight = null;
                        foreach (BoundingPolyline l2 in beamLineVertical[i])
                        {
                            Point3dCollection pTemp = new Point3dCollection();
                            l1.GetLine().IntersectWith(
                                l2.GetLine(), Intersect.ExtendArgument,
                                pTemp, IntPtr.Zero, IntPtr.Zero
                            );
                            if (pTemp.Count == 1)   // 有且仅有一个交点
                            {
                                double distanceLeft = pTemp[0].DistanceTo(l1.GetStartPoint());
                                double distanceRight = pTemp[0].DistanceTo(l1.GetEndPoint());
                                if (distanceLeft < startMLength)
                                {
                                    startMLength = distanceLeft;
                                    lLeft = l2;
                                }
                                if (distanceRight < endMLength)
                                {
                                    endMLength = distanceRight;
                                    lRight = l2;
                                }
                            }
                        }

                        BlockTableRecord btr = trans.GetObject(dwg.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        DimStyleTable dst = trans.GetObject(dwg.DimStyleTableId, OpenMode.ForRead) as DimStyleTable;
                        ObjectId style = dst[Utils.DIM_STYLE_NAME];

                        AlignedDimension rld = new AlignedDimension();
                        rld.DimensionStyle = style;
                        rld.Layer = Utils.DIM_BEAM_NAME;
                        rld.XLine1Point = new Point3d(l1.GetMinPointX(), lLeft.GetMinPointY(), 0);
                        rld.XLine2Point = new Point3d(lLeft.GetMinPointX(), lLeft.GetMinPointY(), 0);
                        rld.DimLinePoint = new Point3d(
                            (l1.GetMinPointX() + lLeft.GetMinPointX()) / 2, 
                            lLeft.GetMinPointY() - 160, 0
                        );

                        btr.AppendEntity(rld);
                        trans.AddNewlyCreatedDBObject(rld, true);

                        AlignedDimension rrd = new AlignedDimension();
                        rrd.DimensionStyle = style;
                        rrd.Layer = Utils.DIM_BEAM_NAME;
                        rrd.XLine1Point = new Point3d(lRight.GetMinPointX(), lRight.GetMinPointY(), 0);
                        rrd.XLine2Point = new Point3d(l1.GetMaxPointX(), lRight.GetMinPointY(), 0);
                        rrd.DimLinePoint = new Point3d(
                            (l1.GetMaxPointX() + lRight.GetMinPointX()) / 2,
                            lRight.GetMinPointY() - 160, 0
                        );

                        btr.AppendEntity(rrd);
                        trans.AddNewlyCreatedDBObject(rrd, true);
                    }
                }
                trans.Commit();
            }   // try
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                ed.WriteMessage("problem due to " + ex.Message);
            }
            finally
            {
                trans.Dispose();
            }

        }   // public void steelBarDim()

        /// <summary>
        /// 显示插件帮助
        /// </summary>
        [CommandMethod("YJKHelp")]
        public void ShowHelp()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            ed.WriteMessage(
                string.Join(
                    "",
                    "\n本插件用于处理YJK软件生成的施工图",
                    "\n  命令示例: ",
                    "\n    YJK - 初始化所有表，计算截面配筋，自动标注支座钢筋",
                    "\n    I   - 非YJK导出DWG初始化，仅计算锚固长度",
                    "\n    H   - 批量修改锚固长度",
                    "\n    O   - 自动生成腰筋标注",
                    "\n    S   - 手动支座钢筋标注",
                    "\n    K   - 计算框选截面的上下配筋",
                    "\n    ShowSteel - 显示选中钢筋的基本信息"
                )
            );

        }   // public void ShowHelp()

    }   // public class YJK

    /// <summary>
    /// 工具类，定义保存精度，正则表达式等
    /// </summary>
    public static class Utils
    {
        public const double DOUBLE_EPS = 1e-4;                              // double判断相等精度
        public const int FIXED_NUM = 4;                                     // double保留小数位数
        public const string ANNO_PATTERN = @"^(\d+)%%132(\d+)$";            // 提取注释正则表达式
        public const string ANNO_PATTERN_J = @"^\((\d+)%%132(\d+)\)$";      // 提取架立筋表达式
        public const string PROFILE_SECTION = @"^(\d+)$";                   // 提取剖切截面标注（梁上）
        public const string PROFILE_ANNO = @"^(\d+)-\d+$";                  // 提取剖切截面标注（截面上）
        public const string BEAM_SECTION_STEEL_NAME = "梁截面纵筋";         // 纵筋图层名
        public const string BEAM_SECTION_GSTEEL_NAME = "梁截面箍筋";
        public const string AXIS_NAME = "轴线";
        public const string BEAM_SECTION_NAME = "梁截面轮廓";
        public const string BEAM_SECTION_ANNO_NAME = "梁截面标注";
        public const string BEAM_SECTION_N_NAME = "梁截面名称";
        public const string DIM_BEAM_NAME = "尺寸标注-梁";
        public const string ANNO_STYLE_NAME = "样式 1";
        public const string DIM_STYLE_NAME = "盈建科结构";
    }

    /// <summary>
    /// 统一Polyline和Line行为
    /// </summary>
    public class BoundingPolyline 
    {
        public bool isPolyline = false;                 // 是否为多段线
        public Line l = null;                           // 存储直线
        public Polyline pl = null;                      // 存储多段线
        public LineSegment3d ls = null;                 // 存储多段线主线段
        public int segmentIndex = -1;                   // 存储多段线主线段的索引
        public bool isOut = false;                      // 是否为外钢筋线
        public int num = -1;                            // 钢筋数量
        public int diameter = -1;                       // 钢筋直线
        public bool isJ = false;                        // 是否为架立筋或不伸入支座钢筋

        /// <summary>
        /// 利用直线构造类
        /// </summary>
        /// <param name="l">CAD类型直线</param>
        public BoundingPolyline(Line l)
        {
            isPolyline = false;
            this.l = l;
        }

        /// <summary>
        /// 重载Equals函数
        /// </summary>
        /// <param name="obj">如果为Entity类型，调用==判断</param>
        public override bool Equals(object obj)
        {
            if (obj.GetType() == typeof(Entity)) return this == (Entity)obj;
            return base.Equals(obj);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        /// 重载==运算符，判断本类与Entity类是否相等
        /// </summary>
        /// <param name="entity1">本类</param>
        /// <param name="entity2">Entity类对象</param>
        /// <returns>首先本类存储类型必须与Entity类相同，且ObjectId相同，则相等，否则不等</returns>
        public static bool operator ==(BoundingPolyline entity1, Entity entity2)
        {
            if (entity1.isPolyline && entity2.GetType() == typeof(Polyline))
            {
                return entity1.pl.Equals(entity2);
            } 
            else if ((!entity1.isPolyline) && entity2.GetType() == typeof(Line))
            {
                return entity1.l.Equals(entity2);
            }
            return false;
        }

        public static bool operator !=(BoundingPolyline entity1, Entity entity2)
        {
            return !(entity1 == entity2);
        }

        /// <summary>
        /// 使用多段线构造本类
        /// </summary>
        /// <param name="pl">CAD类型多段线</param>
        public BoundingPolyline(Polyline pl)
        {
            isPolyline = true;
            double maxLength = 0;
            for (int i = 0; i < pl.NumberOfVertices - 1; i++)
            {
                if (pl.GetLineSegmentAt(i).Length > maxLength)
                {
                    segmentIndex = i;
                    ls = pl.GetLineSegmentAt(i);
                    maxLength = ls.Length;
                }
            }
            this.pl = pl;
            if (ls == null)
                throw new Autodesk.AutoCAD.Runtime.Exception(ErrorStatus.NullObjectId, "没有找到多段线主线段");
        }

        /// <summary>
        /// 使用多段线及是否为梁外钢筋构造本类
        /// </summary>
        /// <param name="pl">CAD类型多段线</param>
        /// <param name="isOut">是否为梁外钢筋</param>
        public BoundingPolyline(Polyline pl, bool isOut) : this(pl)
        {
            this.isOut = isOut;
        }

        /// <summary>
        /// 获取线
        /// </summary>
        /// <returns>多段线或直线</returns>
        public Entity GetLine()
        {
            if (isPolyline) return pl;
            return l;
        }

        /// <summary>
        /// 获取起始点坐标
        /// </summary>
        /// <returns>多段线返回主线段起始点坐标，直线直接返回起始点坐标</returns>
        public Point3d GetStartPoint()
        {
            if (isPolyline) return ls.StartPoint;
            return l.StartPoint;
        }

        /// <summary>
        /// 获取末尾点坐标
        /// </summary>
        /// <returns>多段线返回主线段末尾点坐标，直线直接返回末尾点坐标</returns>
        public Point3d GetEndPoint()
        {
            if (isPolyline) return ls.EndPoint;
            return l.EndPoint;
        }

        /// <summary>
        /// 获取左端点X坐标
        /// </summary>
        /// <returns>左端点X坐标</returns>
        public double GetMinPointX()
        {
            if (isPolyline)
            {
                return Math.Min(ls.StartPoint.X, ls.EndPoint.X);
            }
            return Math.Min(l.StartPoint.X, l.EndPoint.X);
        }

        /// <summary>
        /// 获取右端点X坐标
        /// </summary>
        /// <returns>右端点X坐标</returns>
        public double GetMaxPointX()
        {
            if (isPolyline)
            {
                return Math.Max(ls.StartPoint.X, ls.EndPoint.X);
            }
            return Math.Max(l.StartPoint.X, l.EndPoint.X);
        }

        /// <summary>
        /// 获取下端点Y坐标
        /// </summary>
        /// <returns>下端点Y坐标</returns>
        public double GetMinPointY()
        {
            if (isPolyline)
            {
                return Math.Min(ls.StartPoint.Y, ls.EndPoint.Y);
            }
            return Math.Min(l.StartPoint.Y, l.EndPoint.Y);
        }

        /// <summary>
        /// 获取上端点Y坐标
        /// </summary>
        /// <returns>上端点Y坐标</returns>
        public double GetMaxPointY()
        {
            if (isPolyline)
            {
                return Math.Max(ls.StartPoint.Y, ls.EndPoint.Y);
            }
            return Math.Max(l.StartPoint.Y, l.EndPoint.Y);
        }

        /// <summary>
        /// 获取中点坐标
        /// </summary>
        /// <returns>线段的中点坐标</returns>
        public Point3d GetMidPoint()
        {
            if (isPolyline)
            {
                return ls.StartPoint + (ls.EndPoint - ls.StartPoint) / 2;
            }
            return l.StartPoint + (l.EndPoint - l.StartPoint) / 2; // 中心点坐标
        }

    }   // public class BoundingPolyline 

}   // namespace CAD_YJK
