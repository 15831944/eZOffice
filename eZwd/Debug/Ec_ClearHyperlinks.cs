using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZwd.AddinManager;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.Debug
{
    /// <summary> 清除选择范围内的超链接 </summary>
    [EcDescription("清除选择范围内的超链接")]
    class Ec_ClearHyperlinks : IWordExCommand
    {
        public ExternalCommandResult Execute(Application wdApp, ref string errorMessage, ref object errorObj)
        {
            try
            {
                ClearHyperlinks(wdApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        /// <summary> 生成一个嵌套的域代码： { quote "一九一一年一月{ STYLEREF 1 \s }日" \@"D" }</summary>
        /// <param name="wdApp"></param>
        private static void ClearHyperlinks(Application wdApp)
        {
            wdApp.ScreenUpdating = false;

            var hyperLinks = wdApp.Selection.Range.Hyperlinks;
            for (int i = hyperLinks.Count; i >= 1; i--)
            {
                var hp = hyperLinks[i];
                hp.Delete();
            }

            wdApp.ScreenUpdating = true;
        }

    }
}
