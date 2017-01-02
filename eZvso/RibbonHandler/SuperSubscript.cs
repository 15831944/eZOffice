using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZvso.eZvso_API;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace eZvso.RibbonHandler
{

    /// <summary>
    /// 设置文本的上下标
    /// </summary>
    internal class SuperSubscript
    {
        /// <summary>
        /// 设置文本的上下标
        /// </summary>
        /// <param name="visioApp"></param>
        /// <param name="superOrSubscript">null 表示“正常”，true表示“上标”，false表示“下标”</param>
        public static void SetSuperOrSubScript(Application visioApp, bool? superOrSubscript)
        {
            int undoScopeID1 = visioApp.BeginUndoScope("设置文字上下标");
            try
            {
                SuperSubScript(visioApp, superOrSubscript);
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                MessageBox.Show(errorMessage);
            }
            finally
            {
                visioApp.EndUndoScope(undoScopeID1, true);
            }
        }
        // 开始具体的调试操作
        private static void SuperSubScript(Application vsoApp, bool? superOrSubscript)
        {
            Document doc = vsoApp.ActiveDocument;
            if (doc != null)
            {
                int diagramServices = doc.DiagramServicesEnabled;
                doc.DiagramServicesEnabled = (int)VisDiagramServices.visServiceVersion140 + (int)VisDiagramServices.visServiceVersion150;

                //
                Window win = vsoApp.ActiveWindow;
                Selection sel = vsoApp.ActiveWindow.Selection;

                // 根据不同的编辑或选择情况而进行不同的处理
                int[] selectedShapeIds;
                VisSelectionMode sm = SelectionUtils.GetSelectionMode(sel, out selectedShapeIds);
                switch (sm)
                {
                    case VisSelectionMode.EditingCharactors:
                        var chs = win.SelectedText;
                        if (chs.CharCount > 0)
                        {
                            SuperOrSubScriptCharactors(chs, superOrSubscript);
                        }
                        break;
                    case VisSelectionMode.SingleShape:
                        SuperOrSubScriptShape(sel[1], superOrSubscript);
                        break;
                    case VisSelectionMode.MultiShapes:
                        foreach (Shape shp in sel)
                        {
                            SuperOrSubScriptShape(shp, superOrSubscript);
                        }
                        break;
                }

                // 将 DiagramServicesEnabled 属性复原
                doc.DiagramServicesEnabled = diagramServices;
            }
        }

        /// <summary>
        /// 设置一段字符的上下标
        /// </summary>
        /// <param name="chs"></param>
        /// <param name="superScript">null 表示“正常”，true表示“上标”，false表示“下标”</param>
        private static void SuperOrSubScriptCharactors(Characters chs, bool? superScript)
        {
            if (superScript == null)
            {
                chs.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosNormal;
            }
            else if (superScript.Value)
            {
                chs.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosSuper;
            }
            else
            {
                chs.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosSub;
            }
        }

        /// <summary>
        /// 设置一个形状中所有字符的上下标
        /// </summary>
        /// <param name="shp"></param>
        /// <param name="superScript">null 表示“正常”，true表示“上标”，false表示“下标”</param>
        private static void SuperOrSubScriptShape(Shape shp, bool? superScript)
        {
            List<Shape> nestedShapes = ShapeSearching.GetAllShapes(shp, includeMyself: true);

            foreach (Shape subShape in nestedShapes)
            {
                if (!string.IsNullOrEmpty(subShape.Text))
                {
                    // 如果前期对形状中的字符进行过上下标的单独设置，则这里如果仅仅只设置形状整体的上下标格式，可能会出现形状中部分字符不跟随设置而变化的情况。
                    if (superScript == null)
                    {
                        subShape.Characters.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosNormal;
                        subShape.CellsSRC[(short)VisSectionIndices.visSectionCharacter, (short)0, (short)VisCellIndices.visCharacterPos].FormulaU = VisCellVals.visPosNormal.GetHashCode().ToString();
                    }
                    else if (superScript.Value)
                    {
                        subShape.Characters.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosSuper;
                        subShape.CellsSRC[(short)VisSectionIndices.visSectionCharacter, (short)0, (short)VisCellIndices.visCharacterPos].FormulaU = VisCellVals.visPosSuper.GetHashCode().ToString();
                    }
                    else
                    {
                        subShape.Characters.CharProps[(short)VisCellIndices.visCharacterPos] = (short)VisCellVals.visPosSub;
                        subShape.CellsSRC[(short)VisSectionIndices.visSectionCharacter, (short)0, (short)VisCellIndices.visCharacterPos].FormulaU = VisCellVals.visPosSub.GetHashCode().ToString();
                    }
                }
            }
        }
    }
}
