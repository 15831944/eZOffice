using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZwd.ExternalCommand;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZwd.Debug
{
    class EcTest : IExternalCommand
    {
        public ExternalCommandResult Execute(Microsoft.Office.Interop.Word.Application wdApp, ref string errorMessage, ref object errorObj)
        {
            var rg = wdApp.Selection ;
            if (rg != null)
            {
                MessageBox.Show(rg.Range.Text);
                throw new NullReferenceException(rg.Range.Text);
            }
            return ExternalCommandResult.Succeeded;
        }
    }
}
