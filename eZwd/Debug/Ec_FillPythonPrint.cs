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
    /// <summary> 为 旧版本中Python的Print 方法无括号问题添加括号 </summary>
    [EcDescription("为 旧版本中Python的Print 方法无括号问题添加括号")]
    class Ec_FillPythonPrint : IWordExCommand
    {
        public ExternalCommandResult Execute(Application wdApp, ref string errorMessage, ref object errorObj)
        {
            try
            {

                FillParathesisForPrint(wdApp);
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
        private static void FillParathesisForPrint(Application wdApp)
        {
            wdApp.ScreenUpdating = false;
            Document doc = wdApp.ActiveDocument;

            Selection sel = wdApp.Selection;
            foreach (Paragraph para in sel.Range.Paragraphs)
            {
                var rg = para.Range;
                string t = rg.Text;
                var st = t.Trim();

                if (st.StartsWith("print ") &&
                    (!st.StartsWith("print(") && !st.StartsWith("print (")))
                {
                    // 说明这段字符为： print 'hello world'
                    int i = t.IndexOf("print ", StringComparison.Ordinal);

                    var s2 = t.Insert(i + 6, "(").Insert(i + st.Length + 1, ")");
                    rg.SetRange(rg.Start, rg.End - 1);

                    rg.Text = s2.Substring(0, s2.Length - 1);
                }
            }
            wdApp.ScreenUpdating = true;
        }


    }
}
