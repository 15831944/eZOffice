using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using eZx.AddinManager;
using eZx.AssemblyInfo;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace eZx_AddinManager
{
    internal partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // 将上次插件卸载时保存的程序集数据加载进来

                Excel.Application excelApp = Globals.ThisAddIn.Application;
                form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
                var nodesInfo = AssemblyInfoDllManager.GetInfosFromFile();
                frm.RefreshTreeView(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件加载时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                form_AddinManager frm = form_AddinManager.GetUniqueForm(excelApp);
                var nodesInfo = frm.NodesInfo;
                //
                AssemblyInfoDllManager.SaveAssemblyInfosToFile(nodesInfo);
            }
            catch (Exception ex)
            {
                Debug.Print("AddinManager 插件关闭时出错： \n\r" + ex.Message + "\n\r" + ex.StackTrace);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
