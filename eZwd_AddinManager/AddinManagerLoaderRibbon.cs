using System;
using eZwd.AddinManager;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace eZwd_AddinManager
{
    internal partial class AddinManagerLoaderRibbon
    {
        #region ---   插件的加载与卸载

        private void AddinManagerLoader_Load(object sender, RibbonUIEventArgs e)
        {
            // 将上次插件卸载时保存的程序集数据加载进来
            Application WordApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.InstallAddinManager(WordApp);
        }

        private void AddinManagerLoaderRibbon_Close(object sender, EventArgs e)
        {
            Application WordApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.UninstallAddinManager(WordApp);
        }

        #endregion

        #region ---   点击调试按钮

        private void buttonAddinManager_Click(object sender, RibbonControlEventArgs e)
        {
            Application WordApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.ShowAddinManager(WordApp);
        }

        private void buttonLastCommand_Click(object sender, RibbonControlEventArgs e)
        {
            Application WordApp = Globals.ThisAddIn.Application;
            AddinManagerLoader.LastExternalCommand(WordApp);
        }

        #endregion
    }
}