using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx_API.Entities
{
    public static class ExcelFunction
    {
        #region   ---  工作簿或工作表的匹配

        /// <summary>
        /// 比较两个工作表是否相同。
        /// 判断的标准：先判断工作表的名称是否相同，如果相同，再判断工作表所属的工作簿的路径是否相同，
        /// 如果二者都相同，则认为这两个工作表是同一个工作表
        /// </summary>
        /// <param name="sheet1">进行比较的第1个工作表对象</param>
        /// <param name="sheet2">进行比较的第2个工作表对象</param>
        /// <returns></returns>
        /// <remarks>不用 blnSheetsMatched = sheet1.Equals(sheet2)，是因为这种方法并不能正确地返回True。</remarks>
        public static bool SheetCompare(Worksheet sheet1, Worksheet sheet2)
        {
            bool blnSheetsMatched = false;
            //先比较工作表名称
            if (string.Compare(sheet1.Name, sheet2.Name) == 0)
            {
                Workbook wb1 = sheet1.Parent as Workbook;
                Workbook wb2 = sheet2.Parent as Workbook;
                //再比较工作表所在工作簿的路径
                if (string.Compare(wb1.FullName, wb2.FullName) == 0)
                {
                    blnSheetsMatched = true;
                }
            }
            return blnSheetsMatched;
        }

        /// <summary>
        /// 检测工作簿是否已经在指定的Application程序中打开。
        /// 如果最后此工作簿在程序中被打开（已经打开或者后期打开），则返回对应的Workbook对象，否则返回Nothing。
        /// 这种比较方法的好处是不会额外去打开已经打开过了的工作簿。
        /// </summary>
        /// <param name="wkbkPath">进行检测的工作簿</param>
        /// <param name="Application">检测工作簿所在的Application程序</param>
        /// <param name="blnFileHasBeenOpened">指示此Excel工作簿是否已经在此Application中被打开</param>
        /// <param name="OpenIfNotOpened">如果此Excel工作簿并没有在此Application程序中打开，是否要将其打开。</param>
        /// <param name="OpenByReadOnly">是否以只读方式打开</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static Workbook MatchOpenedWkbk(string wkbkPath, Application Application, ref bool blnFileHasBeenOpened,
            bool OpenIfNotOpened = false, bool OpenByReadOnly = true)
        {
            Workbook wkbk = null;
            if (Application != null)
            {
                //进行具体的检测
                if (File.Exists(wkbkPath)) //此工作簿存在
                {
                    //如果此工作簿已经打开
                    foreach (Workbook WkbkOpened in Application.Workbooks)
                    {
                        if (string.Compare(WkbkOpened.FullName, wkbkPath, true) == 0)
                        {
                            wkbk = WkbkOpened;
                            blnFileHasBeenOpened = true;
                            break;
                        }
                    }

                    //如果此工作簿还没有在主程序中打开，则将此工作簿打开
                    if (!blnFileHasBeenOpened)
                    {
                        if (OpenIfNotOpened)
                        {
                            wkbk = Application.Workbooks.Open(Filename: wkbkPath, UpdateLinks: false,
                                ReadOnly: OpenByReadOnly);
                        }
                    }
                }
            }
            //返回结果
            return wkbk;
        }

        /// <summary>
        /// 检测指定工作簿内是否有指定的工作表，如果存在，则返回对应的工作表对象，否则返回Nothing
        /// </summary>
        /// <param name="wkbk">进行检测的工作簿对象</param>
        /// <param name="sheetName">进行检测的工作表的名称</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static Worksheet MatchWorksheet(Workbook wkbk, string sheetName)
        {
            //工作表是否存在
            Worksheet ValidWorksheet = null;
            foreach (Worksheet sht in wkbk.Worksheets)
            {
                if (string.Compare(sht.Name, sheetName) == 0)
                {
                    ValidWorksheet = sht;
                    return ValidWorksheet;
                }
            }
            //返回检测结果
            return ValidWorksheet;
        }

        #endregion

        #region   ---  几何绘图

        /// <summary>
        /// 将任意形状以指定的值定位在Chart的某一坐标轴中。
        /// </summary>
        /// <param name="ShapeToLocate">要进行定位的形状</param>
        /// <param name="Ax">此形状将要定位的轴</param>
        /// <param name="Value">此形状在Chart中所处的值</param>
        /// <param name="percent">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
        /// 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
        /// 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
        /// <remarks></remarks>
        public static void setPositionInChart(Shape ShapeToLocate, Axis Ax, double Value, double percent = 0)
        {
            Chart cht = (Chart)Ax.Parent;
            if (cht != null)
            {
                //Try          '先考察形状是否是在Chart之中

                //    ShapeToLocate = cht.Shapes.Item(ShapeToLocate.Name)
                //Catch ex As Exception           '如果形状不在Chart中，则将形状复制进Chart，并将原形状删除
                //    ShapeToLocate.Copy()
                //    cht.Paste()
                //    ShapeToLocate.Delete()
                //    ShapeToLocate = cht.Shapes.Item(cht.Shapes.Count)
                //End Try
                //
                switch (Ax.Type)
                {
                    case XlAxisType.xlCategory: //横向X轴
                        double PositionInChartByValue_1 = GetPositionInChartByValue(Ax, Value);
                        Shape with_1 = ShapeToLocate;
                        with_1.Left = (float)(PositionInChartByValue_1 - percent * with_1.Width);
                        break;

                    case XlAxisType.xlValue: //竖向Y轴
                        double PositionInChartByValue = GetPositionInChartByValue(Ax, Value);
                        Shape with_2 = ShapeToLocate;
                        with_2.Top = (float)(PositionInChartByValue - percent * with_2.Height);
                        break;
                    case XlAxisType.xlSeriesAxis:
                        MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                }
            }
        }

        /// <summary>
        /// 将一组形状以指定的值定位在Chart的某一坐标轴中。
        /// </summary>
        /// <param name="ShapesToLocate">要进行定位的形状</param>
        /// <param name="Ax">此形状将要定位的轴</param>
        /// <param name="Values">此形状在Chart中所处的值</param>
        /// <param name="percents">将形状按指定的百分比的宽度或者高度的部位定位到坐标轴的指定值的位置。
        /// 如果其值设定为0，则表示此形状的左端（或上端）定位在设定的位置处，
        /// 如果其值为100，则表示此形状的右端（或下端）定位在设置的位置处。</param>
        /// <remarks></remarks>
        public static void setPositionInChart(Axis Ax, Shape[] ShapesToLocate, double[] Values, double[] Percents = null)
        {
            // ------------------------------------------------------
            //检查输入的数组中的元素个数是否相同
            int Count = ShapesToLocate.Length;
            if (Values.Length != Count)
            {
                MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Percents != null)
            {
                if (Percents.Count() != 1 & Percents.Length != Count)
                {
                    MessageBox.Show("输入数组中的元素个数不相同。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            // ------------------------------------------------------
            Chart cht = (Chart)Ax.Parent;
            //
            double max = Ax.MaximumScale;
            double min = Ax.MinimumScale;
            //
            PlotArea PlotA = cht.PlotArea;
            // ------------------------------------------------------
            Shape shp = default(Shape);
            double Value = 0;
            double Percent = Percents[0];
            double PositionInChartByValue = 0;
            // ------------------------------------------------------

            switch (Ax.Type)
            {
                case XlAxisType.xlCategory: //横向X轴
                    break;


                case XlAxisType.xlValue: //竖向Y轴
                    if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            shp = ShapesToLocate[i];
                            Value = Values[i];
                            if (Percents.Count() > 1)
                            {
                                Percent = Percents[i];
                            }
                            PositionInChartByValue = PlotA.InsideTop + PlotA.InsideHeight * (max - Value) / (max - min);
                            shp.Top = (float)(PositionInChartByValue - Percent * shp.Width);
                        }
                    }
                    else //逆序刻度值，说明Y轴数据为上边小下边大
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            shp = ShapesToLocate[i];
                            Value = Values[i];
                            if (Percents.Count() > 1)
                            {
                                Percent = Percents[i];
                            }
                            PositionInChartByValue = PlotA.InsideTop + PlotA.InsideHeight * (Value - min) / (max - min);
                            shp.Top = (float)(PositionInChartByValue - Percent * shp.Width);
                        }
                    }
                    break;

                case XlAxisType.xlSeriesAxis:
                    MessageBox.Show("暂时不知道这是什么坐标轴", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    break;
            }
        }

        /// <summary>
        /// 根据在坐标轴中的值，来返回这个值在Chart中的几何位置
        /// </summary>
        /// <param name="Ax"></param>
        /// <param name="Value"></param>
        /// <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
        /// 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
        /// <remarks></remarks>
        public static double GetPositionInChartByValue(Axis Ax, double Value)
        {
            double PositionInChartByValue = 0;
            Chart cht = (Chart)Ax.Parent;
            //
            double max = Ax.MaximumScale;
            double min = Ax.MinimumScale;
            //
            PlotArea PlotA = cht.PlotArea;
            switch (Ax.Type)
            {
                case XlAxisType.xlCategory: //横向X轴
                    double PositionInPlot_1 = 0;
                    if (Ax.ReversePlotOrder == false) //正向分类，说明X轴数据为左边小右边大
                    {
                        PositionInPlot_1 = PlotA.InsideWidth * (Value - min) / (max - min);
                    }
                    else //逆序类别，说明X轴数据为左边大右边小
                    {
                        PositionInPlot_1 = PlotA.InsideWidth * (max - Value) / (max - min);
                    }
                    PositionInChartByValue = PlotA.InsideLeft + PositionInPlot_1;
                    break;

                case XlAxisType.xlValue: //竖向Y轴
                    double PositionInPlot = 0;
                    if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                    {
                        PositionInPlot = PlotA.InsideHeight * (max - Value) / (max - min);
                    }
                    else //逆序刻度值，说明Y轴数据为上边小下边大
                    {
                        PositionInPlot = PlotA.InsideHeight * (Value - min) / (max - min);
                    }
                    PositionInChartByValue = PlotA.InsideTop + PositionInPlot;
                    break;
                case XlAxisType.xlSeriesAxis:
                    break;
                    //Debug.Print("暂时不知道这是什么坐标轴")
            }
            return PositionInChartByValue;
        }

        /// <summary>
        /// 根据一组形状在某一坐标轴中的值，来返回这些值在Chart中的几何位置
        /// </summary>
        /// <param name="Ax"></param>
        /// <param name="Values">要在坐标轴中进行定位的多个形状在此坐标轴中的数值</param>
        /// <returns>如果Ax是一个水平X轴，则返回的是坐标轴Ax中的值Value在Chart中的Left值；
        /// 如果Ax是一个竖向Y轴，则返回的是坐标轴Ax中的值Value在Chart中的Top值。</returns>
        /// <remarks></remarks>
        public static double[] GetPositionInChartByValue(Axis Ax, double[] Values)
        {
            int Count = Values.Length;
            double[] PositionInChartByValue = new double[Count - 1 + 1];
            // --------------------------------------------------
            Chart cht = (Chart)Ax.Parent;
            //
            double max = Ax.MaximumScale;
            double min = Ax.MinimumScale;
            double Value = 0;
            //
            PlotArea PlotA = cht.PlotArea;
            switch (Ax.Type)
            {
                case XlAxisType.xlCategory: //横向X轴
                    if (Ax.ReversePlotOrder == false) //正向分类，说明X轴数据为左边小右边大
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            Value = Values[i];
                            PositionInChartByValue[i] = PlotA.InsideLeft + PlotA.InsideWidth * (Value - min) / (max - min);
                        }
                    }
                    else //逆序类别，说明X轴数据为左边大右边小
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            Value = Values[i];
                            PositionInChartByValue[i] = PlotA.InsideLeft + PlotA.InsideWidth * (max - Value) / (max - min);
                        }
                    }
                    break;

                case XlAxisType.xlValue: //竖向Y轴
                    if (Ax.ReversePlotOrder == false) //顺序刻度值，说明Y轴数据为下边小上边大
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            Value = Values[i];
                            PositionInChartByValue[i] = PlotA.InsideTop + PlotA.InsideHeight * (max - Value) / (max - min);
                        }
                    }
                    else //逆序刻度值，说明Y轴数据为上边小下边大
                    {
                        for (UInt16 i = 0; i <= Count - 1; i++)
                        {
                            Value = Values[i];
                            PositionInChartByValue[i] = PlotA.InsideTop + PlotA.InsideHeight * (Value - min) / (max - min);
                        }
                    }
                    break;
                case XlAxisType.xlSeriesAxis:
                    break;
                    //Debug.Print("暂时不知道这是什么坐标轴")
            }
            return PositionInChartByValue;
        }

        #endregion

        #region   ---  Input 对话框操作

        /// <summary> 弹出一个对话框，提示用户输入一个行号 </summary>
        /// <param name="excelApp"></param>
        /// <param name="message"></param>
        /// <returns>如果成功返回一个行号数值，则返回数值，否则返回 null</returns>
        public static int? GetRowNum(Application excelApp, string message)
        {
            dynamic input = excelApp.InputBox(Prompt: message, Type: 8);
            int rowNum = 0;
            if (input is Range && input != null)
            {
                return (input as Range).Row;
            }
            else
            {
                var succ = int.TryParse(input.ToString(), out rowNum);
                if (succ && rowNum > 0)
                {
                    return rowNum;
                }
            }
            return null;
        }


        /// <summary> 弹出一个对话框，提示用户输入多个行号 </summary>
        /// <param name="excelApp"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public static List<int> GetMultipleRowNum(Application excelApp, string message)
        {
            dynamic input = excelApp.InputBox(Prompt: message, Type: 8);
            var rowNums = new List<int>();
            if (input is Range && input != null)
            {
                var rg = input as Range;
                foreach (Range area in rg.Areas)
                {
                    foreach (Range r in area.Rows)
                    {
                        rowNums.Add(r.Row);
                    }
                }
            }
            else
            {
                string txt = input.ToString();
                var nums = txt.Split(',');
                bool succ = false;
                int rowNum;
                foreach (var n in nums)
                {
                    succ = int.TryParse(input.ToString(), out rowNum);
                    if (succ && rowNum > 0)
                    {
                        rowNums.Add(rowNum);
                    }
                }
            }
            return rowNums;
        }
        #endregion

    }
}