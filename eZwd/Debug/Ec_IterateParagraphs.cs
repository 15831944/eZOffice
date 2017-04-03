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
    [EcDescription("对文档中的第一个段落进行迭代，以进入相关操作")]
    class Ec_IterateParagraphs : IWordExCommand
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
            Document doc = wdApp.ActiveDocument;
            foreach (Paragraph p in doc.Paragraphs)
            {
                if (p.Format.FirstLineIndent == 0)
                {
                    var rg = p.Range;
                    if (rg.Tables.Count <= 0 && rg.Text.Length>20)
                    {
                        var res = MessageBox.Show( "\r\n" + rg.Text, "", MessageBoxButtons.OKCancel);
                        if (res == DialogResult.Cancel)
                        {
                            break;
                        }
                    }

                }

            }
        }
    }
}