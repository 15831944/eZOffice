using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            Document doc = visioApp.ActiveDocument;
            if (doc != null)
            {
                MessageBox.Show(doc.Pages.ItemU[1].Name);
                // throw new NullReferenceException(doc.Pages.ItemU[1].Name);
            }
            return ExternalCommandResult.Succeeded;
        }
    }
}
