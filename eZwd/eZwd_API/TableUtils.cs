using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace eZwd.eZwd_API
{
    /// <summary> 与表格相关的操作 </summary>
    public static class TableUtils
    {
        /// <summary> 在文档中插入一个表格 </summary>
        /// <param name="startIndex">表格起始的位置</param>
        /// <param name="data"> 表格所对应的数据，包含表头 </param>
        public static Table InsertTable(Document doc, int startIndex, string[,] data)
        {
            Range rg = doc.Range(Start: startIndex, End: startIndex);
            StringBuilder sb = new StringBuilder();
            int rows = data.GetLength(0);
            int columns = data.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++)
                {
                    sb.Append(data[r, c] + '\r');
                }
            }
            //
            rg.Text = sb.ToString();
            Table tb = rg.ConvertToTable(Separator: WdTableFieldSeparator.wdSeparateByParagraphs,
                NumRows: rows, NumColumns: columns);
            return tb;
        }
    }
}
