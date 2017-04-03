using System;
using System.Collections.Generic;
using DllActivator;
using eZstd.Enumerable;
using eZstd.Mathematics;
using eZx.AddinManager;
using eZx.RibbonHandler;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;

namespace eZx.Debug
{

    [EcDescription("一般性的测试")]
    class EcTest4 : IExcelExCommand
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
        private static void DoSomething(Application excelApp)
        {
            Worksheet sht = excelApp.ActiveSheet;
            Workbook wkbk = excelApp.ActiveWorkbook;
            var rgNum = excelApp.Selection as Range;
            var rgState = rgNum.Offset[0, 3];
            int currentId = 0;
            var numId = new Dictionary<string, int>();
            var rowId = new List<int>();

            foreach (Range c in rgNum)
            {
                int id;
                string num = c.Formula.ToString();
                if (numId.ContainsKey(num))
                {
                    id = numId[num];
                }
                else
                {
                    currentId += 1;
                    numId.Add(num,currentId);
                    id = currentId;
                }
                rowId.Add(id);
            }

            rgState.Value = excelApp.WorksheetFunction.Transpose(rowId.ToArray());
        }

        #endregion
    }
}