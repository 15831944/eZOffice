using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Visio;

namespace eZvso.AddinManager
{
    /// <summary> AddinManager插件 的窗口加载、命令执行 </summary>
    internal class AddinManagerLoader
    {
        #region ---   插件的加载与卸载

        public static void InstallAddinManager(Application VisioApp)
        {
            try
            {
                // 将上次插件卸载时保存的程序集数据加载进来
                form_AddinManager frm = form_AddinManager.GetUniqueForm(VisioApp);
                var nodesInfo = AssemblyInfoDllManager.GetInfosFromSettings();
                frm.RefreshTreeView(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件加载时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }

        public static void UninstallAddinManager(Application VisioApp)
        {
            try
            {
                form_AddinManager frm = form_AddinManager.GetUniqueForm(VisioApp);
                var nodesInfo = frm.NodesInfo;
                //
                // 将窗口中加载的程序集数据保存下来
                AssemblyInfoDllManager.SaveAssemblyInfosToSettings(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件关闭时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }

        #endregion

        #region ---   点击调试按钮

        public static void ShowAddinManager(Application VisioApp)
        {
            form_AddinManager frm = form_AddinManager.GetUniqueForm(VisioApp);
            frm.Show(null);
        }

        public static void LastExternalCommand(Application VisioApp)
        {
            ExCommandExecutor.InvokeCurrentExternalCommand(VisioApp);
        }

        #endregion
    }
}