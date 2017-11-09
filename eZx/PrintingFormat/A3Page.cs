using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace eZx.PrintingFormat
{
    public enum TableType
    {
        /// <summary> 一个WorkSheet中，所有页共享同一个标题与表头 </summary>
        IdenticalTitle,
        /// <summary> 一个WorkSheet中，每一页都有其单独的标题与表头 </summary>
        MultiTitles,
    }

    /// <summary> A3表格页 </summary>
    public class A3Page
    {
        private readonly Worksheet _sht;
        private readonly TableType _tableType;

        public A3Page(Worksheet sht, TableType tableType)
        {
            _sht = sht;
            _tableType = tableType;
        }

        /// <summary> 计算指定页的第一行正文在 Excel 表格中的行号 </summary>
        /// <param name="pageNum"></param>
        /// <returns></returns>
        public int GetContentFirstRowNum(int pageNum)
        {
            switch (_tableType)
            {
                    case TableType.IdenticalTitle:break;
                    case TableType.MultiTitles: break;
            }
            return 2;
        }
    }
}
