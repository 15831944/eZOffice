using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.MarshalReflection;
using eZx.ExternalCommand;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    /// <summary>
    /// 与图表相关的操作
    /// </summary>
    public class ChartHandler
    {
        private readonly Application _app;
        private readonly Chart _chart;

        #region ---   构造函数

        private static ChartHandler _uniqueInstance;

        /// <summary> 获取一个唯一的实例对象 </summary>
        /// <param name="chart"></param>
        /// <returns></returns>
        public static ChartHandler GetUniqueInstance(Chart chart)
        {
            _uniqueInstance = _uniqueInstance ?? new ChartHandler(chart);
            return _uniqueInstance;
        }

        /// <summary> 构造函数 </summary>
        /// <param name="chart"></param>
        private ChartHandler(Chart chart)
        {
            _chart = chart;
            _app = chart.Application;
            //
            _path_desktop = Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "tempData.xlsx");
        }

        #endregion

        #region ---   交换Excel中活动Chart中的每一条数据曲线的X轴与Y轴

        private Chart _lastChart;
        private List<object> _lastX;
        private List<object> _lastY;
        private int _nextExchangeTime;

        /// <summary>
        /// 交换Excel中活动Chart中的每一条数据曲线的X轴与Y轴
        /// </summary>
        public void XYExchange()
        {
            if (_chart != null)
            {
                Series sr = default(Series);
                SeriesCollection src = default(SeriesCollection);
                src = _chart.SeriesCollection();
                if (!_chart.Equals(_lastChart)) //说明是要对一个新的Chart进行操作
                {
                    //
                    _lastX = new List<object>();
                    _lastY = new List<object>();
                    object X;
                    object Y;
                    foreach (Series tempLoopVar_sr in src)
                    {
                        sr = tempLoopVar_sr;
                        X = sr.XValues;
                        Y = sr.Values;
                        //
                        _lastX.Add(X);
                        _lastY.Add(Y);
                        //

                        sr.XValues = Y;
                        sr.Values = X;

                    }
                    _nextExchangeTime = 2;
                }
                else // 说明还是对原来的那个Chart进行操作
                {
                    //此时交换数据时, 使用上一次保存的数据, 而不是直接将现有的Chart中的X与Y交换,
                    //这是因为 : 当X轴为文字，而Y轴为数值时，在交换XY轴后，新的Y轴数据都会变成0，而原来的文字信息在Chart中就不存在了。
                    dynamic X = default(dynamic);
                    object Y = null;
                    for (var i = 1; i <= src.Count; i++)
                    {
                        sr = src.Item(i);
                        X = _lastX[Convert.ToInt32(i - 1)];
                        Y = _lastY[Convert.ToInt32(i - 1)];
                        if (X.Length > 0)
                        {
                            if (_nextExchangeTime % 2 == 0) // 在偶数次交换时，X与Y列使用其原来的数据
                            {
                                sr.XValues = X;
                                sr.Values = Y;
                            }
                            else
                            {
                                sr.XValues = Y;
                                sr.Values = X;
                            }
                        }
                    }
                    _nextExchangeTime++;
                }
                // 将此次操作的Chart中的数据保存起来
                _lastChart = _chart;
            }
            else
            {
                MessageBox.Show("没有找到要进行XY轴交换的图表");
            }
        }
        #endregion

        #region ---   提取图表中的数据

        /// <summary>
        /// 用来临时保存数据的工作簿
        /// </summary>
        /// <remarks>此工作簿用来保存各种临时数据，比如从图表中提取出来的数据情况</remarks>
        private Workbook _tempWkbk;

        /// <summary>
        /// 用来临时保存数据的工作簿的文件路径
        /// </summary>
        /// <remarks>此工作簿位于桌面上的“tempData.xlsx”</remarks>
        private string _path_desktop;


        /// <summary> 提取图表中的数据 </summary>
        /// <remarks></remarks>
        public void ExtractDataFromChart()
        {
            Chart cht = _app.ActiveChart;

            //对Chart中的数据进行提取
            if (cht != null)
            {
                // 打开记录数据的临时工作簿
                if (_tempWkbk == null)
                {
                    if (File.Exists(_path_desktop))
                    {
                        _tempWkbk = (Workbook)Interaction.GetObjectFromFile(_path_desktop);
                        // _tempWkbk = (Workbook)Interaction.GetObject(_path_desktop, null);
                    }
                    else
                    {
                        _tempWkbk = _app.Workbooks.Add();
                        _tempWkbk.SaveAs(_path_desktop);
                    }
                    _tempWkbk.BeforeClose += tempWkbk_BeforeClose;
                }
                //
                Application tempApp = _tempWkbk.Application;
                tempApp.ScreenUpdating = false;


                // 设置写入数据的工作表
                Worksheet sht = _tempWkbk.Worksheets[1]; // 用工作簿中的第一个工作表来存放数据。
                //
                SeriesCollection seriesColl = cht.SeriesCollection();
                Series Chartseries = default(Series);
                //开始提取数据
                short col = (short)1;
                object X = null; // 这里只能将X与Y的数据类型定义为Object，不能是Object()或者Object(,)
                object Y = null;
                string Title = "";
                // 这里不能用For Each Chartseries in SeriesCollection来引用seriesCollection集合中的元素。
                for (var i = 1; i <= seriesColl.Count; i++)
                {
                    // 在VB.NET中，seriesCollection集合中的第一个元素的下标值为1。
                    Chartseries = seriesColl.Item(i);
                    X = Chartseries.XValues;
                    Y = Chartseries.Values;
                    Title = Chartseries.Name;
                    // 将数据存入Excel表中
                    int pointsCount = (X as Array).Length;
                    if (pointsCount > 0)
                    {
                        sht.Cells[1, col].Value = Title;
                        sht.Range[sht.Cells[2, col], sht.Cells[pointsCount + 1, col]].Value =
                            _app.WorksheetFunction.Transpose(X);
                        sht.Range[sht.Cells[2, col + 1], sht.Cells[pointsCount + 1, col + 1]].Value =
                            _app.WorksheetFunction.Transpose(Y);
                        col = (short)(col + 3);
                    }
                }
                sht.Activate();
                _tempWkbk.Save();
                //
                tempApp.Windows[_tempWkbk.Name].Visible = true;
                tempApp.Windows[_tempWkbk.Name].Activate();
                tempApp.ScreenUpdating = true;
                tempApp.Visible = true;
                if (tempApp.WindowState == XlWindowState.xlMinimized)
                {
                    tempApp.WindowState = XlWindowState.xlNormal;
                }
            }
            else
            {
                MessageBox.Show(@"没有找到要进行数据提取的图表");
            }
        }

        private void tempWkbk_BeforeClose(ref bool cancel)
        {
            this._tempWkbk = null;
        }

        #endregion
    }
}
