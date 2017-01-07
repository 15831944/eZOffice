using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZwd.AddinManager;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using eZwd.RibbonHandlers;

namespace eZwd.Debug
{
    /// <summary>
    /// 显示当前光标选择区域的起始处的坐标
    /// </summary>
    [EcDescription("显示选择区域的起止下标值")]
    class Ec_ShowRange : IWordExCommand
    {
        public ExternalCommandResult Execute(Application wdApp, ref string errorMessage, ref object errorObj)
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
            Range rg = wdApp.Selection.Range;
            if (rg != null)
            {
                string t = rg.Text;
                int charactorsCount = t?.Length ?? 0;
                MessageBox.Show($"Start :\t{rg.Start}\r\n End :\t{rg.End}\r\n 字符数 :\t{charactorsCount}");
            }
        }
    }
}