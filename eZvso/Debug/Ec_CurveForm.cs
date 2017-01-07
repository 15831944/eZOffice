using System;
using System.Windows.Forms;
using eZstd.UserControls;
using eZvso.AddinManager;
using eZvso.RibbonHandler.CurveMaker;
using eZvso.AddinManager;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace eZvso.Debug
{
    public class Ec_CurveForm : IVisioExCommand
    {
        public ExternalCommandResult Execute(Application visioApp, ref string errorMessage, ref object errorObj)
        {
            //
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
                frm_CurveParameter f = frm_CurveParameter.GetUniqueInstance(vsoApp);
                //  Form1 f = new Form1();
                f.ShowDialog();
            }
        }
    }
}