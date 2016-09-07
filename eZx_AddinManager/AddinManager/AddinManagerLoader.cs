using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using eZx.AddinManager;
using eZx.AddinManager;
using eZx.ExternalCommand;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace eZx_AddinManager
{
    internal partial class AddinManagerLoader
    {
        private void AddinManagerLoader_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private bool _addinManagerFirstLoaded = true;
        private void buttonAddinManager_Click(object sender, RibbonControlEventArgs e)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
            if (_addinManagerFirstLoaded)
            {
                //// 将上次插件卸载时保存的程序集数据加载进来

                //var nodesInfo = AssemblyInfoDllManager.GetInfosFromFile();
                //frm.RefreshTreeView(nodesInfo);

                //
                _addinManagerFirstLoaded = false;
            }
            else
            {
            }
            frm.Show(null);
        }


        private void buttonLastCommand_Click(object sender, RibbonControlEventArgs e)
        {
            Application excelApp = Globals.ThisAddIn.Application;
            ExternalCommandHandler.InvokeCurrentExternalCommand(excelApp);
        }

    }
}
