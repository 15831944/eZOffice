using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using eZstd.Table;
using eZx.AddinManager;
using eZx.Debug;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    /// <summary> 将工程数量表的工作簿中的多个工程数量表拆分为单独的工作簿 </summary>
    [EcDescription(CommandDescription)]
    public class SepFilesHandler : IExcelExCommand
    {

        #region --- 命令设计

        private const string CommandDescription = @"将工程数量表的工作簿中的多个工程数量表拆分为单独的工作簿";

        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new SepFilesHandler();
            return AddinManagerDebuger.DebugInAddinManager(s.SepFiles,
                excelApp, ref errorMessage, ref errorRange);
        }

        #endregion

        private Application _excelApp;

        /// <summary>
        /// 将工程数量表的工作簿中的多个工程数量表拆分为单独的工作簿
        /// </summary>
        /// <param name="app"></param>
        public ExternalCommandResult SepFiles(Application app)
        {
            _excelApp = app;
            var wkbk = app.ActiveWorkbook; // 要导出的工作簿
            app.ScreenUpdating = false;
       
            // 找出工程数量表
            var qSheets = new List<QuantitySheet>();
            var failedSheets = new List<string>();
            foreach (Worksheet sht in wkbk.Worksheets)
            {
                var Qs = QuantitySheet.GetQuantitySheet(sht);
                if (Qs != null)
                {
                    qSheets.Add(Qs);
                }
                else
                {
                    failedSheets.Add(sht.Name);
                }
            }

            //
            if (qSheets.Count == 0)
            {
                MessageBox.Show(@"未找到任何工程数量表", @"提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return ExternalCommandResult.Cancelled;
            }
            else
            {
                // 搜索模板
                var template = eZstd.Miscellaneous.Utils.ChooseOpenFile("选择Excel模板文件", "Excel 97-2003 工作簿(*.xls)|*.xls", multiselect: false);
                if (template == null)
                {
                    return ExternalCommandResult.Cancelled;
                }

                // 创建文件夹
                var dirPath = Path.Combine(wkbk.Path, "工程量表提取_" + DateTime.Now.ToString("yyyyMMddhhmmss"));
                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }

                // 导出工作表
                foreach (var qs in qSheets)
                {
                    var succ = SeperateSheetIntoNewFile(dirPath, qs, template: template[0]);
                }

                // 提示 并 打开文件夹
                string msg = "所有工作表导出完成 ^_^";
                if (failedSheets.Count != 0)
                {
                    msg = "部分工作表导出完成！\r\n\r\n导出失败的工作表：\r\n";
                    foreach (var fs in failedSheets)
                    {
                        msg += $"{fs}, ";
                    }
                }
                MessageBox.Show(msg, @"提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start(dirPath);
            }
            return ExternalCommandResult.Succeeded;
        }

        /// <summary> 将表格提取到一个全新的工作簿中 </summary>
        /// <param name="qSheet">要提取的工作表</param>
        /// <param name="dirPath">提取出来的工作簿要放在哪一个文件夹中</param>
        /// <returns></returns>
        private  bool SeperateSheetIntoNewFile(string dirPath, QuantitySheet qSheet, string template = null)
        {
            var sht = qSheet.Sheet;
            var fileName = $"{qSheet.Number} {qSheet.Title}.xls";
            var fullName = Path.Combine(dirPath, fileName);

            // 保存文档
            if (!File.Exists(fullName))
            {
                var newWkbk = sht.Application.Workbooks.Add(Template: template);
                sht.Copy(After: newWkbk.Worksheets[newWkbk.Worksheets.Count]);
                newWkbk.SaveAs(Filename: fullName, FileFormat: XlFileFormat.xlExcel8);
                newWkbk.Close();
            }
            else
            {
                var existingWkbk = sht.Application.Workbooks.Open(fullName);
                sht.Copy(After: existingWkbk.Worksheets[existingWkbk.Worksheets.Count]);
                existingWkbk.Close(SaveChanges: true);
            }
            return true;
        }

        /// <summary> 工程数量表 </summary>
        private class QuantitySheet
        {
            public Worksheet Sheet { get; }

            /// <summary> 表格编号 </summary>
            public string Number { get; }

            /// <summary> 表格标题 </summary>
            public string Title { get; }

            public QuantitySheet(Worksheet sheet, string number, string title)
            {
                Sheet = sheet;
                Number = number;
                Title = title;
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="sheet"></param>
            /// <returns>如果不是工程数量表，则返回 null </returns>
            public static QuantitySheet GetQuantitySheet(Worksheet sheet)
            {
                var titleCell = (sheet.Cells[1, 1] as Range).MergeArea;
                var refCell = titleCell.Ex_CornerCell(CornerIndex.BottomRight).Offset[3, 1].MergeArea;
                var instructionCell = refCell.Ex_CornerCell(CornerIndex.UpRight).Offset[0, 1];
                // “图纸参数”那一个单元格。对于一个空表而言，它对应了“C5”单元格
                var paperPara = instructionCell.MergeArea.Ex_CornerCell(CornerIndex.BottomLeft).Offset[1, 0];
                if (paperPara.Value != "图纸参数") return null;
                // 确定了此工作表为工程数量表

                //
                string title = paperPara.Offset[3, 1].Value.ToString();
                title = title.Replace(" ", "");
                //
                string number = paperPara.Offset[5, 1].Value.ToString();
                number = number.Replace("编   号：", "");
                number = number.Replace(" ", "");
                //
                return new QuantitySheet(sheet, number, title);
            }
        }
    }
}