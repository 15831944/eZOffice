using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using eZstd.Enumerable;
using eZstd.Miscellaneous;
using eZwd.AddinManager;
using eZwd.eZwd_API;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.Debug
{
    [EcDescription("测试")]
    internal class EcTest : IWordExCommand
    {
        public ExternalCommandResult Execute(Microsoft.Office.Interop.Word.Application wdApp, ref string errorMessage,
            ref object errorObj)
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
        private void DoSomething(Application wdApp)
        {
            var doc = wdApp.ActiveDocument;
            var data = "这是一段8";
            bool? b = true;

            MessageBox.Show((b == true).ToString());
        }


    }

}