using System;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZstd.Table;
using Microsoft.Office.Interop.Excel;

namespace eZx_API.Entities
{
    public static class ExcelExtension //扩展方法只能在模块中声明。
    {
        /// <summary> Excel的工作表中，最大的行号。在Excel 2016 中，这个值为2^20 </summary>
        internal static int MaxmumRowInExcel = (int)Math.Pow(2, 20);
        /// <summary> Excel的工作表中，最大的列号。在Excel 2016 中，这个值为2^14 </summary>
        internal  static int MaxmumColumnInExcel = (int)Math.Pow(2, 14);

        #region ---   Range

        /// <summary>
        /// 返回Range对象范围中的最右下角点的那个单元格对象
        /// </summary>
        /// <param name="SourceRange">对于对Range.Areas.Item(1)中的单元格区域进行操作</param>
        /// <param name="Corner">要返回哪一个角落的单元格</param>
        /// <returns></returns>
        public static Range Ex_CornerCell(this Range SourceRange, CornerIndex Corner)
        {
            Range myCornerCell = null;
            //
            SourceRange = SourceRange.Areas[1];
            Range LeftTopCell = SourceRange.Cells[1, 1] as Range; ;
            //
            switch (Corner)
            {
                case CornerIndex.BottomRight:
                    myCornerCell = SourceRange.Worksheet.Cells[LeftTopCell.Row + SourceRange.Rows.Count - 1, LeftTopCell.Column + SourceRange.Columns.Count - 1] as Range;
                    break;
                case CornerIndex.UpRight:
                    myCornerCell = SourceRange.Worksheet.Cells[LeftTopCell.Row, LeftTopCell.Column + SourceRange.Columns.Count - 1] as Range; ;
                    break;
                case CornerIndex.BottomLeft:
                    myCornerCell = SourceRange.Worksheet.Cells[LeftTopCell.Row + SourceRange.Rows.Count - 1, LeftTopCell.Column] as Range; ;
                    break;
                case CornerIndex.UpLeft:
                    myCornerCell = LeftTopCell;
                    break;
            }
            return myCornerCell;
        }

        /// <summary>
        /// 收缩Range.Areas.Item(1)的单元格范围到 UsedRange 范围内。
        /// </summary>
        /// <param name="rg">此函数只考虑 rg 所对应的行或列的数量，并不会对 rg 中的单元格的值是否为空进行判断。</param>
        /// <remarks> 
        /// 在选择一个单元格范围时，有时为了界面操作简单，往往会选择一整列或者一整行，但是并不是要对基本所有的单元格进行操作，
        /// 而只需要操作其中有数据的那些区域。此函数即是将选择的整行或者整列的单元格收缩到 UsedRange 范围内。
        /// </remarks>
        public static Range Ex_ShrinkeRange(this Range rg)
        {
            rg = rg.Areas[1];
            int colCount = rg.Columns.Count;
            int rowCount = rg.Rows.Count;
            //
            Range bottomRightCell = rg.Ex_CornerCell(CornerIndex.BottomRight);
            Range usedBottomRightCell = rg.Worksheet.UsedRange.Ex_CornerCell(CornerIndex.BottomRight);

            //  将最下面的单元格收缩到UsedRange的最下面的位置
            if (rowCount == MaxmumRowInExcel && colCount == MaxmumColumnInExcel) // 说明选择了整个表格
            {
                bottomRightCell = usedBottomRightCell;
            }
            else if (rowCount == MaxmumRowInExcel && colCount < MaxmumColumnInExcel) // 说明选择了整列
            {
                bottomRightCell = bottomRightCell.Offset[usedBottomRightCell.Row - MaxmumRowInExcel, 0];
            }
            else if (rowCount < MaxmumRowInExcel && colCount == MaxmumColumnInExcel) // 说明选择了整行
            {
                bottomRightCell = bottomRightCell.Offset[0, usedBottomRightCell.Column - MaxmumColumnInExcel];
                // Else  ' 说明选择了一个有限的范围
            }
            return rg.Worksheet.Range[rg.Cells[1, 1], bottomRightCell];
        }
        
        /// <summary>
        /// 收缩任意一行或者一列，使其最后一个单元格的值不为空
        /// </summary>
        /// <param name="rg">确保此 rg 只包含一行/列（不一定是一整行/列）。</param>
        /// <remarks></remarks>
        public static Range Ex_ShrinkeVectorAndCheckNull(this Range rg)
        {
            bool isRow;  // 指定的 Range 是一行单元格
            if (rg.Rows.Count > 1 && rg.Columns.Count > 1)
            {
                throw new ArgumentException("指定的范围中包含多行或者多列，无法进行收缩。");
            }
            Range ulCell = rg.Ex_CornerCell(CornerIndex.UpLeft);
            int notNullIndex;  // 最后一个 非null 值所对应的行号或者列号
            if (rg.Rows.Count > 1) // 说明选择了一列
            {
                notNullIndex = ulCell.Row;
                isRow = false;
            }
            else if (rg.Columns.Count > 1) // 说明选择了一行
            {
                notNullIndex = ulCell.Column;
                isRow = true;
            }
            else // 说明只选择了一个单元格
            {
                return ulCell;
            }

            //
            rg = rg.Ex_ShrinkeRange();
            // 对每一个单元格进行检测
            object[,] values = rg.Value;
            int checkedIndex = notNullIndex;
            foreach (object v in values)
            {
                if (v != null && !string.IsNullOrEmpty(v.ToString()))
                {
                    notNullIndex = checkedIndex;
                }
                //
                checkedIndex += 1;
            }

            Worksheet sht = rg.Worksheet;
            // 返回校验的一行或者一列的最后一个非空单元格
            Range endCell;
            if (isRow)
            {
                endCell = sht.Cells[ulCell.Row, notNullIndex];
            }
            else
            {
                endCell = sht.Cells[notNullIndex, ulCell.Column];
            }
            return sht.Range[ulCell, endCell];
        }

        /// <summary> 将指定Range的第一个Area的矩形区域进行转置 </summary>
        /// <param name="rg"></param>
        /// <returns></returns>
        public static Range Ex_Transpose(this Range rg)
        {
            rg = rg.Areas.Item[1];
            Range ulCorner = rg.Ex_CornerCell(CornerIndex.UpLeft);
            Range transposedCorner = ulCorner.Offset[rg.Columns.Count - 1, rg.Rows.Count - 1];
            //
            Worksheet sht = rg.Worksheet;
            return sht.Range[ulCorner, transposedCorner];
        }

        #endregion
    }
}
