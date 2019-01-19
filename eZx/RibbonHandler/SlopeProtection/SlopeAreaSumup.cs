using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using eZstd.Enumerable;
using eZstd.Mathematics;
using eZstd.Table;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler.SlopeProtection
{
    internal class SlopeAreaSumup
    {
        private readonly Application _excelApp;

        public SlopeAreaSumup(Application excelApp)
        {
            _excelApp = excelApp;
        }

        // 开始具体的调试操作
        public void Execute(bool containsHeader)
        {
            Range topLeftCell = null;
            var slopes = GetSlopes(containsHeader, ref topLeftCell);

            if (slopes != null && slopes.Count > 0)
            {
                // 1、对横断面数据进行排序
                slopes.Sort(Comparison);

                // 2、 求分段的面积
                var segs = GetArea(slopes);
                // 4、将结果写回 Excel
                var slopesArr = SegmentData<float, double>.ConvertToArr(segs);
                var mileArr = segs.Select(r => $"K{Math.Floor(r.Start / 1000)}+{(r.Start % 1000).ToString("000")}~K{Math.Floor(r.End / 1000)}+{(r.End % 1000).ToString("000")}").ToArray();
                var lengthArr = segs.Select(r => (r.End - r.Start).ToString()).ToArray();
                var lengthArr2 = segs.Select(r => "长度" + (r.End - r.Start).ToString()).ToArray();
                //
                slopesArr = slopesArr.InsertVector<object, object, string>(false, new[] { mileArr, lengthArr },
                    new [] { 1f,1.1f });

                var sht = _excelApp.ActiveSheet;
                RangeValueConverter.FillRange(sht, topLeftCell.Row, topLeftCell.Column + 3, slopesArr, false);
            }
        }

        private int Comparison(KeyValuePair<float, double> mileLength1, KeyValuePair<float, double> mileLength2)
        {
            return mileLength1.Key.CompareTo(mileLength2.Key);
        }

        #region --- 获取数据并进行解析

        /// <summary>
        /// 从 Excel 表格中获取横断面数据
        /// </summary>
        /// <returns></returns>
        public List<KeyValuePair<float, double>> GetSlopes(bool containsHeader, ref Range topLeftCell)
        {
            var slopes = new List<KeyValuePair<float, double>>();
            var rg = _excelApp.Selection as Range;

            if (rg != null)
            {
                // 将选择的范围进行适当的收缩，以匹配有效数据的区域
                Range col = rg.Columns[1];
                Range cell = col.Ex_ShrinkeVectorAndCheckNull().Ex_CornerCell(CornerIndex.BottomLeft);

                // cell 表示选择区域的有效数据区域的最左下角的单元格
                rg = containsHeader
                    ? rg.Rows[$"{2}:{cell.Row - rg.Row + 1}"] // 将表头剃除
                    : rg.Rows[$"{1}:{cell.Row - rg.Row + 1}"];
                topLeftCell = rg.Cells[1, 1];

                object[,] arr = RangeValueConverter.GetRangeValue<object>(rg.Value);
                int errorRow = rg.Row - 1;
                bool keepPromt = true;
                for (int r = 0; r < arr.GetLength(0); r++)
                {
                    errorRow += 1;
                    //
                    float mile = -1;
                    if (!float.TryParse(arr[r, 0].ToString(), out mile))
                    {
                        if (keepPromt)
                        {
                            keepPromt = PrompError($"第{errorRow}行数据有误，无法解析桩号数据。");
                        }
                        continue;
                    }

                    //
                    double slopeLength = 0;
                    if (!double.TryParse(arr[r, 2].ToString(), out slopeLength))
                    {
                        if (keepPromt)
                        {
                            keepPromt = PrompError($"第{errorRow}行数据有误，无法解析边坡长度。");
                        }
                        continue;
                    }
                    //
                    var ml = new KeyValuePair<float, double>(mile, slopeLength);
                    slopes.Add(ml);
                }
            }
            return slopes;
        }

        private bool PrompError(string msg)
        {
            bool keepPrompt = false;
            var res = MessageBox.Show(msg + $"\r\n继续提示?", @"警告",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (res == DialogResult.Yes)
            {
                keepPrompt = true;
            }
            return keepPrompt;
        }

        #endregion

        /// <summary> 根据桩号与对应的斜坡值来计算分段与对应的面积 </summary>
        /// <param name="mile_length">小桩号在前面</param>
        /// <returns></returns>
        public List<SegmentData<float, double>> GetArea(List<KeyValuePair<float, double>> mile_length)
        {
            var res = new List<SegmentData<float, double>>();
            if (mile_length.Count < 2)
            {
                throw new InvalidOperationException("必须指定至少两个桩号才能计算分段面积");
            }
            var lastMl = mile_length[0]; // 上一个桩号
            var startMile = lastMl.Key;
            bool lastIsZero = Math.Abs(lastMl.Value) < 0.0001;

            var area = 0.0; // 分段面积
            for (int i = 1; i < mile_length.Count; i++)
            {
                var ml = mile_length[i];
                var m = ml.Key; // 里程桩号
                var l = ml.Value; // 斜坡长度
                // 求梯形面积
                area += (lastMl.Value + l) * (m - lastMl.Key) / 2;
                //
                var thisIsZero = Math.Abs(l) < 0.0001;
                if (lastIsZero ^ thisIsZero) // 
                {
                    if (thisIsZero) // 说明到了分段的终点
                    {
                        res.Add(new SegmentData<float, double>(startMile, m, area));
                        area = 0;
                    }
                    else // 说明到了分段的起点
                    {
                        startMile = lastMl.Key;
                    }
                }
                else
                {
                }
                lastIsZero = thisIsZero;
                lastMl = ml;
            }
            // 对最后一个桩号进行操作，即最后一个桩号非零的情况下，其面积还没有闭合
            if (!lastIsZero)
            {
                res.Add(new SegmentData<float, double>(startMile, lastMl.Key, area));
            }
            return res;
        }

    }
}