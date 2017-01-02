using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZvso.ExternalCommand;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;
using eZvso.eZvso_API;
namespace eZvso.Debug
{
    public class Ec_Test : IExternalCommand
    {
        public ExternalCommandResult Execute(Application visioApp, ref string errorMessage, ref object errorObj)
        {
            int undoScopeID1 = visioApp.BeginUndoScope("文字属性");
            try
            {
                // SuperScript(visioApp, false);
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