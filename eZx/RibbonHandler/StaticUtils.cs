using System;
using System.Collections.Generic;
using System.Windows.Forms;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler
{
    public static class StaticUtils
    {
        /// <summary>
        /// 对齐打印边界， 以适应图纸的打印区域。
        /// </summary>
        /// <param name="app"></param>
        /// <param name="rightBoundary">右边界的位置，以厘米为单位</param>
        /// <remarks> 由于 单元格的宽度 中一个列宽单位等于“常规”样式中一个字符的宽度。
        /// 对于比例字体，则使用字符 0（零）的宽度，而 形状的宽度和定位 是以 磅为单位，1 point = 1/72 inch；
        /// 而不同的字符宽度略有区别，为了精确定位，可以通过迭代以达到一个差异精度。</remarks>
        public static void 以磅为单位为定位单元格宽度(Application app, double rightBoundary, double bottomBoundary)
        {
            rightBoundary = app.CentimetersToPoints(Centimeters: rightBoundary); // 单位转换
            bottomBoundary = app.CentimetersToPoints(Centimeters: bottomBoundary); // 单位转换
            //
            var sht = app.ActiveSheet as Worksheet;
            var sele = app.Selection as Range;
            if (sele != null)
            {
                var conner = sele.Ex_CornerCell(CornerIndex.BottomRight);

                // 调整行高
                if (bottomBoundary > 0)
                {
                    Range row = sht.Rows[conner.Row];
                    try
                    {
                        //
                        double oldRowHight, newRowHight;
                        double diff = 100;
                        double oldDiff = diff - 1;
                        // 而不同的字符宽度略有区别，为了精确定位，可以通过迭代以达到一个差异精度。
                        while (diff > 0.5 && Math.Abs(diff - oldDiff) > 0.01) //  ' 0.5  ' 单位为磅，即 0.17mm
                        {
                            oldDiff = diff;
                            // 设置列宽 √
                            oldRowHight = row.RowHeight;
                            var ratio = row.Height / oldRowHight; // 磅 转换为 列宽单位 的比例
                            newRowHight = (bottomBoundary - row.Top) / ratio;
                            if (newRowHight < 0)
                            {
                                throw new InvalidOperationException("指定的底边界位于所选单元格上侧！");
                            }
                            row.RowHeight = newRowHight;
                            diff = Math.Abs(newRowHight - oldRowHight);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, @"提示");
                        var line = sht.Shapes.AddLine(0, (float)bottomBoundary, 100, (float)bottomBoundary);
                        line.Placement = XlPlacement.xlFreeFloating;
                    }
                }


                // 调整列宽
                if (rightBoundary > 0)
                {
                    Range col = sht.Columns[conner.Column];
                    try
                    {
                        //
                        double oldcolwidth, newcolwidth;
                        double diff = 100;
                        double oldDiff = diff - 1;
                        // 由于 ColumnWidth 中一个列宽单位等于“常规”样式中一个字符的宽度。对于比例字体，则使用字符 0（零）的宽度，而Width是以 磅为单位，1 point = 1/72 inch；
                        // 而不同的字符宽度略有区别，为了精确定位，可以通过迭代以达到一个差异精度。
                        while (diff > 0.5 && Math.Abs(diff - oldDiff) > 0.01) //  ' 0.00005  ' 单位为磅，即 0.17mm
                        {
                            oldDiff = diff;
                            // 设置列宽 √
                            oldcolwidth = col.ColumnWidth;
                            var ratio = col.Width / oldcolwidth; // 磅 转换为 列宽单位 的比例
                            newcolwidth = (rightBoundary - col.Left) / ratio;
                            if (newcolwidth < 0)
                            {
                                throw new InvalidOperationException("指定的右边界位于所选单元格左侧！");
                            }
                            col.ColumnWidth = newcolwidth;
                            diff = Math.Abs(newcolwidth - oldcolwidth);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, @"提示");
                        var line = sht.Shapes.AddLine((float)rightBoundary, 0, (float)rightBoundary, 100);
                        line.Placement = XlPlacement.xlFreeFloating;
                    }
                }

            }
        }


        /// <summary>
        /// 将桩号数值转换为对应的字符
        /// </summary>
        /// <param name="app"></param>
        /// <param name="selectedRange"></param>
        /// <param name="maxDigits">转换为字符的最大的小数位数</param>
        public static void ConvertStationToString(Application app, Range selectedRange, int maxDigits)
        {

            var firstFillCell = selectedRange.Ex_CornerCell(CornerIndex.UpRight).Offset[0, 1];
            var colCount = selectedRange.Columns.Count;
            if (colCount == 0 || colCount > 2)
            {
                return;
            }
            // 必须只有一列或者两列
            var stationStrings = new List<string>();
            var v = RangeValueConverter.GetRangeValue<object>(selectedRange.Value) as object[,];
            if (v != null)
            {
                var rowsCount = v.GetLength(0);
                var colsCount = v.GetLength(1);
                if (colsCount == 0 || colsCount > 2) return;
                string s1 = null;
                string s2 = null;
                if (colsCount == 1)
                {
                    for (int r = 0; r < rowsCount; r++)
                    {
                        if (v[r, 0] is double)
                        {
                            s1 = GetStationString((double)v[r, 0], maxDigits);
                            stationStrings.Add(s1);
                        }
                        else
                        {
                            stationStrings.Add(null);
                        }
                    }
                }
                else
                {
                    // 共有两列
                    for (int r = 0; r < rowsCount; r++)
                    {
                        if (v[r, 0] is double)
                        {
                            s1 = GetStationString((double)v[r, 0], maxDigits);
                        }
                        else
                        {
                            s1 = null;
                        }
                        if (v[r, 1] is double)
                        {
                            s2 = "~" + GetStationString((double)v[r, 1], maxDigits);
                        }
                        else
                        {
                            s2 = null;
                        }
                        stationStrings.Add(s1 + s2);
                    }
                }
            }
            // 写入到 工作表中
            RangeValueConverter.FillRange(app.ActiveSheet, firstFillCell.Row, firstFillCell.Column, stationStrings.ToArray());
        }

        private static string GetStationString(double station, int maxDigits)
        {
            var res = "";
            var k = (int)Math.Floor(station / 1000);
            var meters = station % 1000;
            var miniMeters = meters % 1;
            if (miniMeters != 0)
            {
                var digits = new string('0', maxDigits);
                res += $"K{k}+{meters.ToString("000." + digits)}";
            }
            else
            {
                // 整米数桩号
                res = $"K{k}+{meters.ToString("000")}";
            }
            return res;
        }

    }
}