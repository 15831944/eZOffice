using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.Geometry;
using eZstd.Miscellaneous;
using eZstd.UserControls;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace eZvso.CurveMaker
{
    public partial class frm_CurveParameter : Form
    {
        private readonly Application _vsoApp;

        #region    ---   构造函数与窗口的打开、关闭

        private static frm_CurveParameter _uniqueInstance;
        /// <summary> 全局唯一的一个窗口实例 </summary>
        /// <param name="vsoApp"></param>
        /// <returns></returns>
        public static frm_CurveParameter GetUniqueInstance(Application vsoApp)
        {
            _uniqueInstance = _uniqueInstance ?? new frm_CurveParameter(vsoApp);
            return _uniqueInstance;
        }

        /// <summary> 构造函数 </summary>
        private frm_CurveParameter(Application vsoApp)
        {
            InitializeComponent();
            this.KeyPreview = true;
            this.FormClosing += OnFormClosing;
            this.KeyDown += OnKeyDown;
            //
            _vsoApp = vsoApp;
            //
            textBoxTolerance.Text = @"0.01";
            textBoxTolerance.PositiveOnly = true;
            textBox_degree.Text = @"2";
            textBox_degree.PositiveOnly = true;
            textBox_degree.IntegerOnly = true;
            //
            ConstructDatagridview(dataGridView1 as eZDataGridView);
            // 事件绑定
            radioButton_spline.CheckedChanged += RadioButtonSplineOnCheckedChanged;
        }

        /// <summary> 按下 ESC 时关闭窗口 </summary>
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
        }

        #endregion

        #region    ---   DataGridView

        private BindingList<LocationPoint> _points;
        private void ConstructDatagridview(eZDataGridView eZDgv)
        {
            // 设置表格信息
            eZDgv.KeyDelete = true;
            eZDgv.ManipulateRows = true;
            eZDgv.ShowRowNumber = true;
            eZDgv.SupportPaste = true;
            //
            eZDgv.AllowUserToAddRows = true;
            //
            Column_X.ValueType = typeof(double);
            Column_X.DataPropertyName = "X";

            //
            Column_Y.ValueType = typeof(double);
            Column_Y.DataPropertyName = "Y";
            //

            //
            _points = new BindingList<LocationPoint>
            {
                AllowNew = true
            };
            _points.AddingNew += PointsOnAddingNew;

            // 事件关联
            eZDgv.DataError += EZDgvOnDataError;
            eZDgv.DataSource = _points;
        }

        #endregion

        #region    ---   事件处理



        private void RadioButtonSplineOnCheckedChanged(object sender, EventArgs eventArgs)
        {
            panelTolerance.Visible = radioButton_spline.Checked;
        }

        private void PointsOnAddingNew(object sender, AddingNewEventArgs e)
        {
            e.NewObject = new LocationPoint();
        }

        private void EZDgvOnDataError(object sender, DataGridViewDataErrorEventArgs dataGridViewDataErrorEventArgs)
        {
            MessageBox.Show(@"坐标点的值必须为数值",
                @"格式错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        #endregion

        #region    ---   绘图

        private void buttonOk_Click(object sender, EventArgs e)
        {
            if (radioButton_spline.Checked && _points.Count > 1)
            {
                // 距离容差越大，生成的曲线与控制点的偏差越大，曲线越光滑
                double tol = textBoxTolerance.ValueNumber;
                Drawer.DrawSplineOnPage(_vsoApp.ActivePage, _points.ToList(), tol);
            }
            else if (radioButton_polyline.Checked && _points.Count > 1)
            {
                Drawer.DrawPolylineOnPage(_vsoApp.ActivePage, _points.ToList());
            }
            else if (radioButton_bezier.Checked && _points.Count > 1)
            {
                int degree = (int)textBox_degree.ValueNumber;
                Drawer.DrawBezierOnPage(_vsoApp.ActivePage, _points.ToList(), degree: degree);
            }
            else if (radioButton_nurbs.Checked && _points.Count > 1)
            {

            }

        }
        #endregion

    }


}
