using System;
using DllActivator;
using eZx.AddinManager;
using eZx.RibbonHandler;
using eZx.RibbonHandler.SlopeProtection;
using Microsoft.Office.Interop.Excel;

namespace eZx.Debug
{
    [EcDescription("边坡断面信息处理")]
    class EcSlopeInfo : IExcelExCommand
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
        private static void DoSomething(Application excelApp)
        {
            Worksheet sht = excelApp.ActiveSheet;
            Workbook wkbk = excelApp.ActiveWorkbook;
            //
            var slpHdl = new SlopeInfoHandler(excelApp);
            slpHdl.Execute(containsHeader:true);
        }

        #endregion
    }
}