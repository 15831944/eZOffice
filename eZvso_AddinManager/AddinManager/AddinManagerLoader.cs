using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using eZvso.AddinManager;
using eZvso.ExternalCommand;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;

namespace eZvso.AddinManager
{
    internal class AddinManagerLoader
    {
        #region ---   插件的加载与卸载

        public static void InstallAddinManager(Application excelApp)
        {
            try
            {
                // 将上次插件卸载时保存的程序集数据加载进来
                form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
                var nodesInfo = AssemblyInfoDllManager.GetInfosFromFile();
                frm.RefreshTreeView(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件加载时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }

        public static void UninstallAddinManager(Application excelApp)
        {
            try
            {
                form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
                var nodesInfo = frm.NodesInfo;
                //
                // 将窗口中加载的程序集数据保存下来
                AssemblyInfoDllManager.SaveAssemblyInfosToFile(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件关闭时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }
        #endregion

        #region ---   点击调试按钮

        public static void ShowAddinManager(Application excelApp)
        {
            form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
            frm.Show(null);
        }

        public static void LastExternalCommand(Application excelApp)
        {
            ExternalCommandHandler.InvokeCurrentExternalCommand(excelApp);
        }

        #endregion
    }
}
