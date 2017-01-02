using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;

namespace eZvso.eZvso_API
{
    /// <summary>
    /// 与Visio界面中的选择集相关的操作
    /// </summary>
    public static class SelectionUtils
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sel"></param>
        /// <param name="selectedShapeIds">在 <see cref="VisSelectMode.visSelModeSkipSuper"/> 模式下筛选得到的选择形状的ID集合</param>
        /// <returns></returns>
        public static VisSelectionMode GetSelectionMode(Selection sel, out int[] selectedShapeIds)
        {
            VisSelectionMode sm;
            sel.IterationMode = (int)VisSelectMode.visSelModeSkipSuper; // 默认情况下，IterationMode 为 visSelModeSkipSuper + visSelModeSkipSub
            Array ids;
            sel.GetIDs(out ids);
            selectedShapeIds = ids as int[];
            // 1. 
            if (selectedShapeIds == null || selectedShapeIds.Length <= 0) { return VisSelectionMode.NoSelection; }

            // 2. 
            if (selectedShapeIds.Length == 1)
            {
                Shape shp = sel[1];
                return shp.IsOpenForTextEdit ? VisSelectionMode.EditingCharactors : VisSelectionMode.SingleShape;
            }
            else
            {
                return VisSelectionMode.MultiShapes;
            }
        }


    }

    /// <summary>
    /// Visio当前界面中在进行何种选择操作
    /// </summary>
    public enum VisSelectionMode
    {
        /// <summary> 未选择任何对象 </summary>
        NoSelection = 0,
        /// <summary> 正在编辑某个形状中的文字，此时只可能选中一个形状 </summary>
        EditingCharactors,
        /// <summary> 选择单个形状，而且此形状并未进入文字编辑状态，即其< </summary>
        SingleShape,
        /// <summary> 选择多个形状 </summary>
        MultiShapes
    }
}
