using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZvso.AddinManager;
using eZvso.eZvso_API;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
namespace eZvso.Debug
{
    [EcDescription("删除数值字符")]
    public class Ec_Test : IVisioExCommand
    {
        public ExternalCommandResult Execute(Application visioApp, ref string errorMessage, ref object errorObj)
        {
            int undoScopeID1 = visioApp.BeginUndoScope("文字属性");
            try
            {
                visioApp.ShowChanges = false;
                visioApp.ScreenUpdating = 0;
                DoSomething(visioApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
            finally
            {
                visioApp.ShowChanges = true;
                visioApp.ScreenUpdating = 1;
                visioApp.EndUndoScope(undoScopeID1, true);
            }
        }


        // 开始具体的调试操作
        private static void DoSomething(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            if (doc != null)
            {
                var p = vsoApp.ActivePage;
                var shps = ShapeSearching.GetAllShapes(p);
                foreach (Shape shp in shps)
                {
                    MessageBox.Show(shp.ID.ToString());
                }
            }
        }

    }
}