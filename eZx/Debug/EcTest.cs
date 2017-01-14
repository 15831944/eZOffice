using System;
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

        }


        #endregion
    }
}