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
    [EcDescription("对于有很多行数据的工程量表，自动将多数据行进行分隔，并插入小计行")]
    class EcInsertSumupRow : IExcelExCommand
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
        }

        #region ---   具体的调试操作

        // 开始具体的调试操作
        private static void DoSomething(Application app)
        {
            Worksheet sht = app.ActiveSheet;
            Workbook wkbk = app.ActiveWorkbook;


            var rg = app.InputBox("选择第二张表格中的区域（包括第一个表格中的小计行）", Type: 8) as Range;
            if (rg != null)
            {
                int? lastRow = 0;
                lastRow = ExcelFunction.GetRowNum(app, "最后一行数据的行号：");
                if (lastRow != null)
                {
                    var sumupRow = rg.Rows[1] as Range;
                    Range indexColumn = sht.Range[rg.Cells[2, 1], rg.Cells[rg.Rows.Count - 1, 1]] as Range;

                    var startRow = sumupRow.Row + 1;
                    var dataRowsCount = indexColumn.Count;
                    //
                    SumRowHandler.InsertSumupRow(app, sumupRow: sumupRow, indexColumn: indexColumn,
                        startRow: startRow, dataRowsCount: dataRowsCount, lastRow: lastRow.Value);
                }
                else
                {
                    MessageBox.Show(@"请输入一个数值");
                }
            }
            return;
            //
            SumRowHandler.InsertSumupRow(app, sht.Range["A36:Q36"], sht.Range["A37:A64"], 37, 28, 1010 + 7);
            return;
            var slpHdl = new SlopeInfoHandler(app);
            slpHdl.Execute(containsHeader: true);
        }
        

        #endregion
    }
}