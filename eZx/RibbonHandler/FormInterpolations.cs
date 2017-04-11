using System;
using System.Windows.Forms;
using eZstd.Enumerable;
using eZstd.Mathematics;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    public partial class FormInterpolations : Form
    {
        private Application _exApp;

        #region ---   构造函数与窗口开启关闭

        private static FormInterpolations _uniqueInstance;

        /// <summary> 获取全局唯一窗口实例 </summary>
        public static FormInterpolations GetUniqueInstance(Application excelApp)
        {
            _uniqueInstance = _uniqueInstance ?? new FormInterpolations(excelApp);
            return _uniqueInstance;
        }

        /// <summary> 构造函数
        /// </summary> <param name="excelApp"></param>
        private FormInterpolations(Application excelApp)
        {
            InitializeComponent();
            //
            _exApp = excelApp;
            rangeSource.SetApplication(excelApp);
            rangeGetorI.SetApplication(excelApp);
            rangeGetorD.SetApplication(excelApp);
            //
            KeyPreview = true;
        }

        private void FormInterpolations_FormClosing(object sender, FormClosingEventArgs e)
        {
            Hide();
            e.Cancel = true;
        }

        #endregion

        #region ---   界面事件

        private void FormInterpolations_KeyDown(object sender, KeyEventArgs e)
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

        private void buttonOk_Click(object sender, EventArgs e)
        {
            try
            {
                double[] srcX, srcY, interpX;

                if (GetSplineSrc(out srcX, out srcY, out interpX))
                {
                    double[] interpY = null;
                    if (radioButton_Spline.Checked)
                    {
                        interpY = Interpolation.SplineInterpolation(srcX, srcY, interpX);
                    }
                    else if (radioButton_whatever.Checked)
                    {
                        interpY = Interpolation.LinearInterpolation(srcX, srcY, interpX);
                    }
                    if (interpY != null)
                    {
                        // 将结果写入 Excel 表格中
                        Range destCell;
                        if (rangeGetorD.Range == null)
                        {
                            Range rg = destCell = rangeGetorI.Range.Cells[1];
                            destCell = rg.Offset[0, 1];
                        }
                        else
                        {
                            destCell = rangeGetorD.Range.Cells[1];
                        }
                        Worksheet sht = _exApp.ActiveSheet;
                        RangeValueConverter.FillRange(sht, destCell.Row, destCell.Column, interpY, true);
                    }
                }
                else
                {
                    MessageBox.Show(@"无法找到有效的数据源", @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + "\r\n", // + ex.StackTrace
                    @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region --- 样条插值

        private bool GetSplineSrc(out double[] srcX, out double[] srcY, out double[] interpX)
        {
            srcX = null;
            srcY = null;
            interpX = null;

            if ((rangeSource.RangeX != null) && (rangeSource.RangeY != null) && (rangeGetorI.Range != null))
            {
                srcX = getFirstColumnData(rangeSource.RangeX);
                srcY = getFirstColumnData(rangeSource.RangeY);
                interpX = getFirstColumnData(rangeGetorI.Range);
                if (srcX.Length != srcY.Length)
                {
                    return false;
                }
                return true;
            }
            return false;
        }

        private double[] getFirstColumnData(Range rg)
        {
            Range c = rg.Columns[1];
            var d = RangeValueConverter.GetRangeValue<double>(rg.Value, false, 0);
            var cv = ArrayConstructor.GetColumn(d, 0);
            return cv;
        }

        #endregion
    }
}