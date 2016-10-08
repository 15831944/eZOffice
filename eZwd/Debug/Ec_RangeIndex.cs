using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZwd.ExternalCommand;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.Debug
{
    /// <summary> 查看选择的区域的Range的范围 </summary>
    class Ec_RangeIndex : IExternalCommand
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
                errorMessage = ex.Message + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }


        // 开始具体的调试操作
        private static void DoSomething(Application wdApp)
        {
            var sel = wdApp.Selection;
            MessageBox.Show($"start: {sel.Start} \r\nend: {sel.End}");
        }
    }
}
