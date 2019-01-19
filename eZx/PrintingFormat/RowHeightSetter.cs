using System;
using System.Windows.Forms;
using eZx.AddinManager;
using eZx.Debug;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.PrintingFormat
{
    /// <summary> A3页面打印设置 </summary>
    [EcDescription(CommandDescription)]
    class RowHeightSetter : IExcelExCommand
    {
        #region --- 命令设计

        private const string CommandDescription = @"设置 A3 表格的打印格式";

        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new A3PageSetup();
            return AddinManagerDebuger.DebugInAddinManager(s.SetupA3Page,
                excelApp, ref errorMessage, ref errorRange);
        }

        #endregion

        private Application _excelApp;

        /// <summary> </summary>
        public ExternalCommandResult SetContentRowHeight(Application excelApp)
        {
            _excelApp = excelApp;
            excelApp.ScreenUpdating = false;
            var sht = excelApp.ActiveSheet as Worksheet;
            var rg = excelApp.Selection as Range;
            if (rg != null)
            {
                SetRowHeight(excelApp.Selection as Range);
            }
            return ExternalCommandResult.Succeeded;
        }

        private void SetRowHeight(Range rg)
        {
            var cell = rg.Cells[1, 1] as Range;
            var titleRowNum = cell.Row;
            var sht = rg.Worksheet;
            // 标题行
            sht.Rows[titleRowNum].RowHeight = 30;
            // 
            for (int r = titleRowNum + 1; r < titleRowNum + 6; r++)
            {
                sht.Rows[r].RowHeight = 20;
            }
            // 列序号所在行
            sht.Rows[titleRowNum + 6].RowHeight = 15;
            // 正文所在行
            var firstContentRowNum = titleRowNum + 7;
            var lastContentRowNum = firstContentRowNum + 28;

            for (int r = firstContentRowNum; r <= lastContentRowNum; r++)
            {
                sht.Rows[r].RowHeight = 20;
            }
            // 选择表格所在的行
            sht.Rows[$"{titleRowNum}:{lastContentRowNum}"].Select();
        }
    }
}