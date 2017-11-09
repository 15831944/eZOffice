using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.Enumerable;
using eZstd.Miscellaneous;
using eZx.AddinManager;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.Debug.工程量清单
{
    [EcDescription(CommandDescription)]
    class 规范数据源 : IExcelExCommand
    {
        #region --- 命令设计
        private const string CommandDescription = @"工程量清单规范化";
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new 规范数据源();
            return AddinManagerDebuger.DebugInAddinManager(s.QuantityList,
                excelApp, ref errorMessage, ref errorRange);
        }
        #endregion

        private static Item[] ItemSources;

        /// <summary> </summary>
        public ExternalCommandResult QuantityList(Application excelApp)
        {
            excelApp.ScreenUpdating = false;
            int rowNum = 0;
            try
            {
                Range rg = excelApp.Selection as Range;
                var shtDes = excelApp.ActiveSheet as Worksheet;
                var wkbkDes = excelApp.ActiveWorkbook;
                Range dest = (rg.Cells[1] as Range).Offset[0, 8];
                //// 规范数据源
                //var wkbk =
                //    excelApp.Workbooks.Open(@"C:\Users\Administrator\Desktop\徐敏 工程量清单\格式.xlsx",
                //        ReadOnly: false);

                Worksheet shtsrc = wkbkDes.Sheets["格式列表"];
                ItemSources = GetStandardSource(shtsrc.UsedRange).ToArray();
                // 要匹配的数据
                var arrV = RangeValueConverter.GetRangeValue<object>(rg.Value) as object[,];
                var items = new List<Item>();
                for (int r = 0; r < arrV.GetLength(0); r++)
                {
                    rowNum = dest.Row + r ;
                    var unit = GetString(arrV[r, 3]);
                    if (unit == null || unit == @"/")
                    {
                        unit = null;
                    }
                    var item = new Item(GetString(arrV[r, 0]), GetString(arrV[r, 1]), GetString(arrV[r, 2]), unit);
                    // 比较与匹配
                    var m = ItemSources.FirstOrDefault(rr => rr.子目号 == item.子目号);
                    if (m != null)
                    {
                        var m2 = ItemSources.FirstOrDefault(rr => rr.子目名称 == item.子目名称);
                        if (m2 != null)
                        {
                            m.备注 = Matched.匹配;
                        }
                        else
                        {
                            m.备注 = Matched.未匹配;

                        }
                        items.Add(m);
                    }
                    else
                    {
                        item.备注 = Matched.未包含;
                        items.Add(new Item("", "", "", "") { 备注 = Matched.未包含 });
                    }
                }
                // 写入
                //
                var arr = new List<object[]>();
                foreach (var i in items)
                {
                    arr.Add(i.ToArray());
                }
                //
                var arr2 = ArrayConstructor.FromList2D(arr);

                RangeValueConverter.FillRange(shtDes, dest.Row, dest.Column, arr2);
                wkbkDes.Activate();
                shtDes.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("出错行：" + rowNum.ToString() + ex.Message + ex.StackTrace);
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
            return ExternalCommandResult.Succeeded;
        }

        public List<Item> GetStandardSource(Range rg)
        {
            var items = new List<Item>();
            var src = RangeValueConverter.GetRangeValue<object>(rg.Value) as object[,];
            for (int r = 1; r < src.GetLength(0); r++)
            {
                var unit = GetString(src[r, 3]);
                if (unit == null || unit == @"/")
                {
                    unit = null;
                }
                var item = new Item(GetString(src[r, 0]), GetString(src[r, 1]), GetString(src[r, 2]), unit);
                items.Add(item);
            }
            return items;
        }

        private string GetString(object obj)
        {
            if (obj == null)
            {
                return null;
            }
            else
            {
                return obj.ToString();
            }
        }
    }
}