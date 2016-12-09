using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace eZx
{
    public partial class Ribbon_eZx
    {
        private HelpLocation frmHelpLocation;
        private HelpLocationSettings settings1 = new HelpLocationSettings();

        public void Group_Help_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            if (frmHelpLocation == null)
            {
                frmHelpLocation = new HelpLocation();
            }
            frmHelpLocation.ShowDialog();
        }

        #region 加载文档

        // 打开帮助文档所在文件夹
        public void btn_OfficeHelp_Click(object sender, RibbonControlEventArgs e)
        {
            settings1.Reload();
            string DirePath = settings1.OfficeHelp;
            //
            if (Directory.Exists(DirePath))
            {
                Process.Start(DirePath);
            }
            else
            {
                MessageBox.Show(@"指定的帮助文档不存在，请重新设置帮助文档路径。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 打开帮助文档
        public void btn_ExcelHelp_Click(object sender, RibbonControlEventArgs e)
        {
            settings1.Reload();
            string filePath = settings1.ExcelHelp;
            //
            if (File.Exists(filePath))
            {
                try
                {
                    Process.Start(filePath);
                }
                catch (Exception)
                {
                    MessageBox.Show(@"指定的帮助文档无法打开。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(@"指定的帮助文档不存在，请重新设置帮助文档路径。", "出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}