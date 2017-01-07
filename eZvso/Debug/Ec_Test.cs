using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZvso.AddinManager;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
namespace eZvso.Debug
{
    [EcDescription("一般性的测试")]
    public class Ec_Test : IVisioExCommand
    {
        public ExternalCommandResult Execute(Application visioApp, ref string errorMessage, ref object errorObj)
        {
            int undoScopeID1 = visioApp.BeginUndoScope("文字属性");
            try
            {
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
            finally
            {
                visioApp.EndUndoScope(undoScopeID1, true);
            }
        }


    }
}