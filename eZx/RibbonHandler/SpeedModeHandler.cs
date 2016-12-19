using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.Enumerable;
using eZstd.MarshalReflection;
using eZstd.Mathematics;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    internal class SpeedModeHandler
    {

        public void Main(Application excelApp)
        {
            double[] xValues;
            double[] yValues;


            // 获取原始数据
            Worksheet sht = excelApp.ActiveSheet;

            //
            Range valueRange = RangeValueConverter.GetRange(sht, 1, 1, 117, 2);
            var value = valueRange.Value;

            double[] vDate = ArrayConstructor.GetColumn<double>(RangeValueConverter.GetRangeValue<double>(value, false, 0), 0);

            double[,] vValue = RangeValueConverter.GetRangeValue<double>(value, false, 1);
            // 其他计算参数
            int type = (int)sht.Cells[1, 4].Value;
            int newCount = (int)sht.Cells[2, 4].Value;

            // var sp = new SpeedMode(XdateNum.Select(r => (double)r).ToArray(), Y);
            var sp = new SpeedMode(vDate, ArrayConstructor.GetColumn(vValue, 0));
            SpeedMode.ShrinkResult res;
            switch (type)
            {
                case 1:
                    {
                        res = sp.ShrinkByIdAverage(newCount);
                        if (res == SpeedMode.ShrinkResult.Succeed)
                        {
                            // 绘制数据
                            RangeValueConverter.FillRange(sht, 1, 5, sp.GetX(), colPrior: true);
                            RangeValueConverter.FillRange(sht, 1, 6, sp.GetY(), colPrior: true);
                        }
                        break;
                    }
                case 2:
                    {
                        res = sp.ShrinkByXAxis(newCount);
                        if (res == SpeedMode.ShrinkResult.Succeed)
                        {
                            // 绘制数据

                            RangeValueConverter.FillRange(sht, 1, 7, sp.GetX(), colPrior: true);
                            RangeValueConverter.FillRange(sht, 1, 8, sp.GetY(), colPrior: true);
                        }
                        else
                        {
                            MessageBox.Show(sp.ErrorMessage);
                        }
                        break;
                    }
            }
        }

        #region ---   ShrinkByPointCount

        /// <summary>
        /// 以只考虑数据点个数的方式进行缩减
        /// </summary>
        /// <param name="srcX">数据源中的X数据</param>
        /// <param name="srcY">数据源中的Y数据</param>
        /// <param name="newCount">缩减后的新数据点个数</param>
        /// <param name="srcD">要将缩减后的数据放置在哪里，此属性中只包含一个单元格，表示整个缩减后的曲线的左上角单元格</param>
        public static void ShrinkByPointCount(Range srcX, Range srcY, int newCount, Range srcD)
        {
            double[] x = GetColumnData(srcX);
            double[] y = GetColumnData(srcY);
            if (x.Length > 2 && x.Length == y.Length)
            {
                //
                var sp = new SpeedMode(x, y);
                SpeedMode.ShrinkResult res = sp.ShrinkByIdAverage(newCount);
                if (res == SpeedMode.ShrinkResult.Succeed)
                {
                    Worksheet sht = srcD.Worksheet;
                    // 绘制数据
                    RangeValueConverter.FillRange(sht, srcD.Row, srcD.Column, sp.GetX(), colPrior: true);
                    RangeValueConverter.FillRange(sht, srcD.Row, srcD.Column + 1, sp.GetY(), colPrior: true);
                }
            }
            else
            {
                throw new ArgumentException(@"X与Y数据点的个数必须相同而且至少为2。");
            }
        }

        #endregion

        #region ---   ShrinkByXRange

        /// <summary>
        /// 以只考虑曲线的X轴数据区间的分段的方式进行缩减
        /// </summary>
        /// <param name="srcX">数据源中的X数据</param>
        /// <param name="srcY">数据源中的Y数据</param>
        /// <param name="xSeg">要将X轴所占区间分为多少段</param>
        /// <param name="srcD">要将缩减后的数据放置在哪里，此属性中只包含一个单元格，表示整个缩减后的曲线的左上角单元格</param>
        public static void ShrinkByXRange(Range srcX, Range srcY, int xSeg, Range srcD)
        {
            double[] x = GetColumnData(srcX);
            double[] y = GetColumnData(srcY);
            if (x.Length > 2 && x.Length == y.Length)
            {
                //
                var sp = new SpeedMode(x, y);
                SpeedMode.ShrinkResult res = sp.ShrinkByXAxis(xSeg);
                if (res == SpeedMode.ShrinkResult.Succeed)
                {
                    Worksheet sht = srcD.Worksheet;
                    // 绘制数据
                    RangeValueConverter.FillRange(sht, srcD.Row, srcD.Column, sp.GetX(), colPrior: true);
                    RangeValueConverter.FillRange(sht, srcD.Row, srcD.Column + 1, sp.GetY(), colPrior: true);
                }
            }
            else
            {
                throw new ArgumentException(@"X与Y数据点的个数必须相同而且至少为2。");
            }
        }

        #endregion

        #region ---   子方法

        /// <summary>
        /// 提取一列数据中有效的数值或者时间数据
        /// </summary>
        /// <param name="columnOrRow">一列或者一行</param>
        /// <returns></returns>
        private static double[] GetColumnData(Range columnOrRow)
        {
            columnOrRow = columnOrRow.Ex_ShrinkeRange();
            List<double> data = new List<double>();
            double num;
            DateTime dt;
            foreach (object v in columnOrRow.Value)
            {
                if (v == null)
                {
                    continue;
                }
                if (v is double)
                {
                    data.Add((double)v);
                }
                else if (v is DateTime)
                {
                    data.Add(((DateTime)v).ToOADate());
                }
                else if (double.TryParse(v.ToString(), out num))
                {
                    data.Add(num);
                }
                else if (DateTime.TryParse(v.ToString(), out dt))
                {
                    data.Add(dt.ToOADate());
                }
            }
            return data.ToArray();
        }

        #endregion


    }
}