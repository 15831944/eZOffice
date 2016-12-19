using System;
using DllActivator;
using eZx.AddinManager;
using eZx.RibbonHandler;
using Microsoft.Office.Interop.Excel;

namespace eZx.Debug
{
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
                errorMessage = ex.Message + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        // 开始具体的调试操作
        private static void DoSomething(Application excelApp)
        {
            FormSpeedModeHandler f = FormSpeedModeHandler.GetUniqueInstance(excelApp);
            f.Show(null);
        }

        #region ---   具体的调试操作

        #endregion
    }
}