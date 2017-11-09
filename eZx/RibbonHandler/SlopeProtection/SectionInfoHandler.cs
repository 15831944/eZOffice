using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.RibbonHandler.SlopeProtection
{
    public class SlopeInfoHandler
    {
        private readonly Application _excelApp;
        public SortedSet<MileageInfo> Slopes;

        public SlopeInfoHandler(Application excelApp)
        {
            _excelApp = excelApp;
        }

        public void Execute(bool containsHeader)
        {
            Range topLeftCell = null;
            var slopes = GetSlopes(true, ref topLeftCell);
            if (slopes != null && slopes.Count > 0)
            {
                // 1、对横断面数据进行排序，排序后集合中不会有重复的里程
                slopes = Sort_SumDuplicate(slopes);

                // 2、 去掉有测量值的定位断面
                var sortedSections = new Stack<MileageInfo>();
                double duplicateMileage = double.Epsilon;
                for (int i = slopes.Count - 1; i >= 0; i--)
                {
                    var slp = slopes[i];
                    if (slp.Mileage == duplicateMileage)
                    {
                        // 说明出现重复里程，此时从众多重复里程中仅保留有测量值的那个里程断面
                        if (slp.Type == MileageInfoType.Measured)
                        {
                            // 替换掉原集合中的值
                            sortedSections.Peek().Override(slp);
                        }
                    }
                    else
                    {
                        // 说明与上一个桩号不重复
                        sortedSections.Push(slp);
                        duplicateMileage = slp.Mileage;
                    }
                }

                // 3、 进行插值，此时 sortedSlopes 中，较小的里程位于堆的上面，
                // 而且集合中只有“定位”与“测量”两种断面，并没有“插值”断面
                var allSections = new List<MileageInfo>();
                var count = sortedSections.Count;
                var smallestSection = sortedSections.Pop();
                allSections.Add(smallestSection);
                //
                var lastSec = smallestSection;
                var lastType = smallestSection.Type;
                for (int i = 1; i < count; i++)
                {
                    var slp = sortedSections.Pop();

                    if (slp.Type != lastType)
                    {
                        // 说明从定位断面转到了测量断面，或者从测量断面转到了定位断面，此时要进行断面插值
                        var interpMile = (lastSec.Mileage + slp.Mileage)/2;
                        var interpSec = new MileageInfo(interpMile, MileageInfoType.Interpolated, 0);
                        allSections.Add(interpSec);
                    }
                    //
                    allSections.Add(slp);

                    lastSec = slp;
                    lastType = slp.Type;
                }
                // 4、将结果写回 Excel
                var slopesArr = MileageInfo.ConvertToArr(allSections);
                var sht = _excelApp.ActiveSheet;
                ;
                RangeValueConverter.FillRange(sht, topLeftCell.Row, topLeftCell.Column + 3, slopesArr, false);
            }
        }


        private List<MileageInfo> Sort_SumDuplicate(List<MileageInfo> slopes)
        {
            slopes.Sort(new MileageCompare());
            // 排序后集合中可能会有重复的里程(比如一个里程中，将一种边坡防护分两级坡分开算)

            // 现在将排序后相同里程的数据进行相加
            var distinctSlopes = new List<MileageInfo>();
            MileageInfo lastSlope = slopes[0];
            double lastMile = lastSlope.Mileage;
            distinctSlopes.Add(lastSlope);
            for (int i = 1; i < slopes.Count; i++)
            {
                var sp = slopes[i];
                if (sp.Mileage == lastMile)
                {
                    lastSlope.SpLength += sp.SpLength;
                }
                else
                {
                    lastSlope = sp;
                    lastMile = sp.Mileage;
                    distinctSlopes.Add(lastSlope);
                }
            }

            return distinctSlopes;
        }
        #region --- 获取数据并进行解析

        /// <summary>
        /// 从 Excel 表格中获取横断面数据
        /// </summary>
        /// <returns></returns>
        public List<MileageInfo> GetSlopes(bool containsHeader, ref Range topLeftCell)
        {
            var slopes = new List<MileageInfo>();
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
                    MileageInfoType tp;
                    var tpStr = arr[r, 1].ToString();
                    if (MileageInfo.TypeMapping.ContainsKey(tpStr))
                    {
                        tp = MileageInfo.TypeMapping[tpStr];
                    }
                    else
                    {
                        if (keepPromt)
                        {
                            keepPromt = PrompError($"第{errorRow}行数据有误，无法解析横断面类型。");
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
                    var s = new MileageInfo(mileage: mile, type: tp, slopeLength: slopeLength);
                    slopes.Add(s);
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
    }
}