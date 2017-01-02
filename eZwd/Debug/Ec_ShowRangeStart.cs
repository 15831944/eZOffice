using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZwd.ExternalCommand;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using eZwd.RibbonHandlers;

namespace eZwd.Debug
{
    /// <summary>
    /// 显示当前光标选择区域的起始处的坐标
    /// </summary>
    class Ec_ShowRangeStart : IExternalCommand
    {
        public ExternalCommandResult Execute(Microsoft.Office.Interop.Word.Application wdApp, ref string errorMessage, ref object errorObj)
        {
            try
            {
                DoSomething(wdApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        // 开始具体的调试操作
        private static void DoSomething(Application wdApp)
        {
            MessageBox.Show(wdApp.Selection.Start.ToString());
        }
    }
}