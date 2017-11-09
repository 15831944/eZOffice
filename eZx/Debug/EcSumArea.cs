using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DllActivator;
using eZstd.Enumerable;
using eZstd.Mathematics;
using eZx.AddinManager;
using eZx.RibbonHandler;
using eZx.RibbonHandler.SlopeProtection;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.Debug
{

    [EcDescription("根据断面与坡长统计面积")]
    class EcSumArea : IExcelExCommand
    {
        private Application _excelApp;

        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            DllActivator_eZx dat = new DllActivator_eZx();
            dat.ActivateReferences();
            try
            {
                _excelApp = excelApp;
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
        private void DoSomething(Application excelApp)
        {
            var slpHdl = new SlopeAreaSumup(excelApp);
            slpHdl.Execute(containsHeader:true);
        }
        
        #endregion
    }
}