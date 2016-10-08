using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using eZx.ExternalCommand;
using eZx.RibbonHandler;
using Microsoft.Office.Interop.Excel;
using eZstd.MarshalReflection;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.Debug
{
    class EcTest : IExternalCommand
    {
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
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
            string _path_desktop = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "tempData.xlsx");

            string progId = "Excel.Application";

            // 开始具体的调试操作
            var ch = ChartHandler.GetUniqueInstance(excelApp.ActiveChart);
            ch.ExtractDataFromChart();
        }
    }
}