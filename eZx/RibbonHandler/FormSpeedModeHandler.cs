using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    public partial class FormSpeedModeHandler : Form
    {
        #region ---   Fields

        private Application _app;
        #endregion

        #region ---   构造函数与窗口开启关闭

        private static FormSpeedModeHandler _uniqueInstance;

        /// <summary> 获取全局唯一窗口实例 </summary>
        public static FormSpeedModeHandler GetUniqueInstance(Application excelApp)
        {
            _uniqueInstance = _uniqueInstance ?? new FormSpeedModeHandler(excelApp);
            return _uniqueInstance;
        }
        private FormSpeedModeHandler(Application excelApp)
        {
            InitializeComponent();
            _app = excelApp;
            rangeSource.SetApplication(excelApp);
            rangeGetorD.SetApplication(excelApp);
        }

        private void FormSpeedModeHandler_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            e.Cancel = true;
        }
        #endregion

        #region ---   界面事件

        private void FormSpeedModeHandler_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {

            Close();
        }

        #endregion

        private void button_Ok_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton_PointCount.Checked)
                {
                    int newCount = (int)numericUpDown_PointsSegments.Value;
                    if (rangeSource.RangeX != null && rangeSource.RangeY != null && rangeGetorD.Range != null)
                    {
                        SpeedModeHandler.ShrinkByPointCount(rangeSource.RangeX, rangeSource.RangeY, newCount, rangeGetorD.Range);
                    }
                }
                else if (radioButton_XSegment.Checked)
                {
                    int xSeg = (int)numericUpDown_PointsSegments.Value;
                    if (rangeSource.RangeX != null && rangeSource.RangeY != null && rangeGetorD.Range != null)
                    {
                        SpeedModeHandler.ShrinkByXRange(rangeSource.RangeX, rangeSource.RangeY, xSeg, rangeGetorD.Range);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        }
}
