using System;
using eZvso.AddinManager;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;

namespace eZvso_AddinManager
{
    internal partial class AddinManagerLoaderRibbon
    {
        #region ---   插件的加载与卸载

        private void AddinManagerLoader_Load(object sender, RibbonUIEventArgs e)
        {
            // 将上次插件卸载时保存的程序集数据加载进来
            Application VisioApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.InstallAddinManager(VisioApp);
        }

        private void AddinManagerLoaderRibbon_Close(object sender, EventArgs e)
        {
            Application VisioApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.UninstallAddinManager(VisioApp);
        }

        #endregion

        #region ---   点击调试按钮

        private void buttonAddinManager_Click(object sender, RibbonControlEventArgs e)
        {
            Application VisioApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.ShowAddinManager(VisioApp);
        }

        private void buttonLastCommand_Click(object sender, RibbonControlEventArgs e)
        {
            Application VisioApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.LastExternalCommand(VisioApp);
        }

        #endregion
    }
}