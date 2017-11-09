using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eZstd.Enumerable;
using eZstd.Miscellaneous;
using eZx.AddinManager;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;

namespace eZx.Debug.工程量清单
{
    [EcDescription(CommandDescription)]
    class 规范编号 : IExcelExCommand
    {
        #region --- 命令设计
        private const string CommandDescription = @"工程量清单规范化";
        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new 规范编号();
            return AddinManagerDebuger.DebugInAddinManager(s.QuantityList,
                excelApp, ref errorMessage, ref errorRange);
        }
        #endregion

        /// <summary> </summary>
        public ExternalCommandResult QuantityList(Application excelApp)
        {
            var sele = excelApp.Selection as Range;
            sele = sele.Ex_ShrinkeRange();
            var sht = excelApp.ActiveSheet;
            var v = RangeValueConverter.GetRangeValue<object>(sele.Value, false, 0) as object[,];

            List<string> comb = new List<string>();
            string cellV;
            string lastP = v[0, 0].ToString();
            for (int r = 0; r < v.GetLength(0); r++)
            {
                if (v[r, 0] != null)
                {
                    cellV = v[r, 0].ToString();
                    cellV.Trim();
                    if (cellV.StartsWith("-"))
                    {
                        comb.Add(lastP + cellV);
                    }
                    else
                    {
                        comb.Add(cellV);
                        lastP = cellV;
                    }
                }
                else
                {
                    comb.Add(null);
                    lastP = null;
                }
            }

            //
            var desti = sele.Offset[0, 1].Cells[1] as Range;
            var arr = comb.ToArray();
            RangeValueConverter.FillRange(sht, desti.Row, desti.Column, arr);


            return ExternalCommandResult.Succeeded;
        }

        public List<Item> GetStandardSource(Worksheet sht)
        {
            var rg = sht.Range[""];
            var items = new List<Item>();
            var src = RangeValueConverter.GetRangeValue<object>(rg.Value) as object[,];
            for (int r = 0; r < src.GetLength(0); r++)
            {
                var unit = GetString(src[r, 4]);
                if (unit == null || unit == @"/")
                {
                    unit = null;
                }
                var item = new Item(GetString(src[r, 0]), GetString(src[r, 2]), GetString(src[r, 3]), unit);
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