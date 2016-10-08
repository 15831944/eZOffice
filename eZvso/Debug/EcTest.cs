using System;
using System.Windows.Forms;
using eZvso.ExternalCommand;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace eZvso.Debug
{
    class EcTest : IExternalCommand
    {
        public ExternalCommandResult Execute(Application visioApp, ref string errorMessage, ref object errorObj)
        {
            try
            {
                DoSomething(visioApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        // 开始具体的调试操作
        private static void DoSomething(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            if (doc != null)
            {
                MessageBox.Show(doc.Pages.ItemU[1].Name);
                // throw new NullReferenceException(doc.Pages.ItemU[1].Name);
            }
        }
    }
}