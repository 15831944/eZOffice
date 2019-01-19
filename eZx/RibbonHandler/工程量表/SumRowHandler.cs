using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace eZx.RibbonHandler
{
    /// <summary>
    /// 工程量表中有多页数据时，对小计行的插入或合并等操作
    /// </summary>
    public class SumRowHandler
    {
        /// <summary>
        /// 对于有很多行数据的工程量表，自动将多数据行进行分隔，并插入小计行
        /// </summary>
        /// <param name="sumupRow">小计行</param>
        /// <param name="indexColumn">每一个表中的序号列，其不包括小计行的单元格</param>
        /// <param name="startRow">第二张表中的第一行数据所在的行号</param>
        /// <param name="dataRowsCount">每一张表中的数据行数，不包括小计行</param>
        /// <param name="lastRow">所有数据的最后一行的行号</param>
        public static void InsertSumupRow(Application app, Range sumupRow, Range indexColumn, int startRow,
            int dataRowsCount, int lastRow)
        {
            app.ScreenUpdating = false;
            var sht = sumupRow.Worksheet;
            Range pasteDataRow = null;
            Range pasteIndexRg = null;
            // 一张表格中所有的数据行数（包括小计行）
            int tableDataRowsCount = dataRowsCount + 1;

            try
            {
                var sumRowNum = startRow + dataRowsCount;

                while (sumRowNum - lastRow <= dataRowsCount)
                {
                    // 插入一个小计行
                    var insertRow = sht.Rows[sumRowNum] as Range;
                    insertRow.Insert(Shift: XlInsertShiftDirection.xlShiftDown,
                        CopyOrigin: XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                    // 复制小计行数据
                    pasteDataRow = sumupRow.Offset[sumRowNum - startRow + 1, 0];
                    sumupRow.Copy(pasteDataRow);

                    // 复制序号
                    pasteIndexRg = indexColumn.Offset[sumRowNum - startRow + 1 - tableDataRowsCount, 0];
                    indexColumn.Copy(pasteIndexRg);

                    //
                    sumRowNum += tableDataRowsCount;
                    lastRow += 1;
                }
            }
            catch (Exception)
            {
                //
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// 将同一Sheet中的多个工程量表进行合并，即删除小计行，并将多个表格中的数据合并
        /// </summary>
        /// <param name="app"></param>
        /// <param name="page1">要进行小计行删除第一页数据，不包括表头，但是包括最后一个小计行</param>
        /// <param name="sumRows">第一页数据中，要进行删除的多个小计行在Excel中的行号</param>
        /// <param name="lastRow">所有要处理的最后一行数据的行号，<paramref name="page1"/>的首行到<paramref name="lastRow"/>之间的所有数据会进行小计行删除操作</param>
        public static void DeleteSumupRow(Application app, Range page1, List<int> sumRows, int lastRow)
        {
            app.ScreenUpdating = false;
            var sht = page1.Worksheet;
            //

            try
            {
                var firstRow = page1.Row;
                var pageRowCount = page1.Rows.Count;
                //
                var deleteRows = new List<int>();
                foreach (var sr in sumRows)
                {
                    var relativeRow = sr - firstRow + 1;
                    if (relativeRow <= pageRowCount)
                    {
                        deleteRows.Add(relativeRow);
                    }
                }
                // 将数值大的排在前面
                deleteRows.Sort(CompareInt);
                var firstRowPage = firstRow;
                var deleteRowNum = 0;
                Range deleteRow;
                int tableCount = (int)Math.Floor((lastRow - firstRow + 1) / (double)pageRowCount);
                for (int t = 0; t < tableCount; t++)
                {
                    firstRowPage = firstRow + t * (pageRowCount - deleteRows.Count);
                    foreach (var dr in deleteRows)
                    {
                        deleteRowNum = firstRowPage + dr - 1;
                        deleteRow = sht.Rows[deleteRowNum];
                        deleteRow.Delete(Shift:XlDeleteShiftDirection.xlShiftUp);
                    }
                }
            }
            catch (Exception)
            {
                //
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private static int CompareInt(int v1, int v2)
        {
            return v2.CompareTo(v1);
        }
    }
}