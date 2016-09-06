using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZx.ExternalCommand;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.ExternalCommandTest
{
    class ExternalCommandTest : IExternalCommand
    {
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref object errorObj)
        {
            Range rg = excelApp.Selection as Range;
            if (rg != null)
            {
                MessageBox.Show(rg.Address);
            }
            return ExternalCommandResult.Succeeded;
        }
    }
}
