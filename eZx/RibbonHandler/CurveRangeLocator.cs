using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    /// <summary> 对二维曲线数据进行操作时，用来定位对应的数据源 </summary>
    public partial class CurveRangeLocator : UserControl
    {
        #region --- 数据源

        public readonly Range RangeX;
        public readonly Range RangeY;
        public readonly Range RangeZ;

        #endregion

        private Application _excelApp;
        /// <summary> 为控件设置一个 Application 对象，此方法必须在构造函数执行后立即执行。 </summary>
        public void SetApplication(Application excelApp)
        {
            if (_excelApp == null)  // 只能设置一次
            {
                if (excelApp != null)
                {
                    _excelApp = excelApp;
                    //
                    rangeGetorX.SetApplication(excelApp);
                    rangeGetorY.SetApplication(excelApp);
                    rangeGetorD.SetApplication(excelApp);
                }
                else
                {
                    throw new NullReferenceException("the excel application object can not be null.");
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public CurveRangeLocator()
        {
            InitializeComponent();

            //
            RangeX = rangeGetorX.Range;
            RangeY = rangeGetorX.Range;
            RangeZ = rangeGetorX.Range;
        }

        /// <summary> 选择数据源或者目标数据的单元格 </summary>
        /// <param name="sender"><see cref="System.Windows.Forms.Button"/>对象</param>
        private void button_srcXY_Click(object sender, EventArgs e)
        {

            Range tagRg = button_srcXY.Tag as Range;
            var inputResult = _excelApp.InputBox(
                Prompt: "选择初始曲线的单元格",
                Title: "选择单元格区域",
                Default: (tagRg != null) ? tagRg.Address : "A1",
                Type: 8);
            if (!(inputResult is Range)) return;

            // 对不同的按钮设置不同的
            Range rg = inputResult as Range;
            button_srcXY.Tag = rg;
            // 从XY中拆解出X与Y这两列数据
            Range sourceX;
            Range sourceY;
            if (SeperateXY(rg, out sourceX, out sourceY))
            {
                rangeGetorX.SetRange(sourceX, isOuterEvent: false, raisePossibleEvent: true);
                rangeGetorY.SetRange(sourceY, isOuterEvent: false, raisePossibleEvent: true);
                //
                button_srcXY.Tag = _excelApp.Union(sourceX, sourceY);
            }
        }

        /// <summary>
        /// 将选择的XY数据源拆分为X与Y
        /// </summary>
        /// <param name="sourceRange"></param>
        /// <param name="sourceX"></param>
        /// <param name="sourceY"></param>
        /// <returns>如果拆解成功，则返回true</returns>
        private static bool SeperateXY(Range sourceRange, out Range sourceX, out Range sourceY)
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

        private void rangeGetorX_RangeChanged(object sender, Range e)
        {
            if (e == null)
            {
                button_srcXY.Tag = rangeGetorY.Range;
            }
            else
            {
                Range r = e.Columns[1];
                rangeGetorX.SetRange(r, isOuterEvent: true, raisePossibleEvent: false);

                button_srcXY.Tag = (rangeGetorY.Range == null)
                    ? r
                    : _excelApp.Union(r, rangeGetorY.Range);
            }


        }

        private void rangeGetorY_RangeChanged(object sender, Range e)
        {
            if (e == null)
            {
                button_srcXY.Tag = rangeGetorX.Range;
            }
            else
            {
                Range r = e.Columns[1];
                rangeGetorY.SetRange(r, isOuterEvent: true, raisePossibleEvent: false);

                button_srcXY.Tag = (rangeGetorX.Range == null)
                    ? r
                    : _excelApp.Union(rangeGetorX.Range, r);
            }

        }
    }
}
