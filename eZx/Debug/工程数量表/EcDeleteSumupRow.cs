using System;
using System.Windows.Forms;
using DllActivator;
using eZx.AddinManager;
using eZx.RibbonHandler;
using eZx.RibbonHandler.SlopeProtection;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.Debug
{
    [EcDescription("将同一Sheet中的多个工程量表进行合并，即删除小计行，并将多个表格中的数据合并")]
    class EcDeleteSumupRow : IExcelExCommand
    {
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            DllActivator_eZx dat = new DllActivator_eZx();
            dat.ActivateReferences();
            try
            {
                DoSomething(excelApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }

        #region ---   具体的调试操作

        // 开始具体的调试操作
        private static void DoSomething(Application app)
        {
            Worksheet sht = app.ActiveSheet;
            Workbook wkbk = app.ActiveWorkbook;


            var rg = app.InputBox("选择要进行小计行删除的第一页数据，包括小计行。", Type: 8) as Range;
            if (rg != null)
            {
                var sumRows = ExcelFunction.GetMultipleRowNum(app, "选择第一页中的多个小计行：");
                if (sumRows.Count == 0) return;
                //
                int? lastRow = ExcelFunction.GetRowNum(app, "选择要处理的最后一行数据的行号：");
                if (lastRow == null) return;
                //
                SumRowHandler.DeleteSumupRow(app, page1: rg, sumRows: sumRows, lastRow: lastRow.Value);
            }
        }


        #endregion
    }
}