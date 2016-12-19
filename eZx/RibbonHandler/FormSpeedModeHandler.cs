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

        /// <summary> 数据源中的X数据 </summary>
        private Range _srcX;
        /// <summary> 数据源中的Y数据 </summary>
        private Range _srcY;
        /// <summary> 要将缩减后的数据放置在哪里，此属性中只包含一个单元格，表示整个缩减后的曲线的左上角单元格 </summary>
        private Range _srcD;

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

        private void button_Ok_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton_PointCount.Checked)
                {
                    int newCount = (int)textBoxNum_PointsSegments.ValueNumber;
                    if (_srcX != null && _srcY != null && _srcD != null)
                    {
                        SpeedModeHandler.ShrinkByPointCount(_srcX, _srcY, newCount, _srcD);
                    }
                }
                else if (radioButton_XSegment.Checked)
                {
                    int xSeg = (int)textBoxNum_PointsSegments.ValueNumber;
                    if (_srcX != null && _srcY != null && _srcD != null)
                    {
                        SpeedModeHandler.ShrinkByXRange(_srcX, _srcY, xSeg, _srcD);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace, @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region ---   选择数据源

        /// <summary> 选择数据源或者目标数据的单元格 </summary>
        /// <param name="sender"><see cref="System.Windows.Forms.Button"/>对象</param>
        private void button_Datasource_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = sender as System.Windows.Forms.Button;
            if (btn != null)
            {
                Range tagRg = btn.Tag as Range;
                var inputResult = _app.InputBox(
                    Prompt: "选择初始曲线的单元格",
                    Title: "选择单元格区域",
                    Default: (tagRg != null) ? tagRg.Address : "A1",
                    Type: 8);
                if (!(inputResult is Range)) return;

                // 对不同的按钮设置不同的
                Range rg = inputResult as Range;
                btn.Tag = rg;
                switch (btn.Name)
                {
                    case "button_srcX":
                        _srcX = rg.Columns[1];
                        btn.Tag = _srcX;
                        textBox_srcX.Text = _srcX.Address;
                        button_srcXY.Tag = (_srcY == null) ? _srcX : _app.Union(_srcX, _srcY);
                        break;
                    case "button_srcY":
                        _srcY = rg.Columns[1];
                        btn.Tag = _srcY;
                        textBox_srcY.Text = _srcY.Address;
                        button_srcXY.Tag = (_srcX == null) ? _srcY : _app.Union(_srcX, _srcY);
                        break;
                    case "button_srcD":
                        _srcD = rg.Cells[1, 1];
                        btn.Tag = _srcD;
                        textBox_srcD.Text = _srcD.Address;
                        break;
                    case "button_srcXY":
                        // 从XY中拆解出X与Y这两列数据
                        Range sourceX;
                        Range sourceY;
                        if (SeperateXY(rg, out sourceX, out sourceY))
                        {
                            _srcX = sourceX;
                            button_srcX.Tag = _srcX;
                            textBox_srcX.Text = _srcX.Address;
                            //
                            _srcY = sourceY;
                            button_srcY.Tag = _srcY;
                            textBox_srcY.Text = _srcY.Address;
                            //
                            btn.Tag = _app.Union(sourceX, sourceY);
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// 将选择的XY数据源拆分为X与Y
        /// </summary>
        /// <param name="sourceRange"></param>
        /// <param name="sourceX"></param>
        /// <param name="sourceY"></param>
        /// <returns>如果拆解成功，则返回true</returns>
        private bool SeperateXY(Range sourceRange, out Range sourceX, out Range sourceY)
        {
            if (sourceRange.Areas.Count >= 2)
            {
                sourceX = sourceRange.Areas[1].Columns[1];
                sourceY = sourceRange.Areas[2].Columns[1];
                return true;
            }
            else
            {
                if (sourceRange.Columns.Count >= 2)
                {
                    sourceX = sourceRange.Columns[1];
                    sourceY = sourceRange.Columns[2];
                    return true;
                }
                else
                {
                    sourceX = null;
                    sourceY = null;
                    return false;
                }
            }
        }


        #endregion

        private void textBox_Datasource_Enter(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox text = sender as System.Windows.Forms.TextBox;
            if (text != null)
            {
                Range rg = null;
                switch (text.Name)
                {
                    case "textBox_srcX": rg = button_srcX.Tag as Range; break;
                    case "textBox_srcY": rg = button_srcY.Tag as Range; break;
                    case "textBox_srcD": rg = button_srcD.Tag as Range; break;
                }
                if (rg != null)
                {
                    rg.Select();
                }
            }
        }
    }
}
