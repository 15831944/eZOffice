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
        #region   ---  列号的字符与数值的转换

        /// <summary>
        /// 将Excel表中的列的数值编号转换为对应的字符
        /// </summary>
        /// <param name="ColNum">Excel中指定列的数值序号</param>
        /// <returns>以字母序号的形式返回指定列的列号</returns>
        /// <remarks>1对应A；26对应Z；27对应AA</remarks>
        public static string ConvertColumnNumberToString(int ColNum)
        {
            // 关键算法就是：连续除以基，直至商为0，从低到高记录余数！
            // 其中value必须是十进制表示的数值
            //intLetterIndex的位数为digits=fix(log(value)/log(26))+1
            //本来算法很简单，但是要解决一个问题：当value为26时，其26进制数为[1 0]，这样的话，
            //以此为下标索引其下面的strAlphabetic时就会出错，因为下标0无法索引。实际上，这种特殊情况下，应该让所得的结果成为[26]，才能索引到字母Z。
            //处理的方法就是，当所得余数remain为零时，就将其改为26，然后将对应的商的值减1.
            if (ColNum < 1)
            {
                MessageBox.Show("列数不能小于1");
            }
            List<byte> intLetterIndex = new List<byte>();
            //
            int quotient = 0; //商
            byte remain = 0; //余数
            //
            byte i = (byte)0;
            do
            {
                quotient = (int)Math.Floor((double)ColNum / 26); // (int)(Conversion.Fix((double)ColNum / 26)); //商
                remain = (byte)(ColNum % 26); //余数
                if (remain == 0)
                {
                    intLetterIndex.Add(26);
                    quotient--;
                }
                else
                {
                    intLetterIndex.Add(remain);
                }
                i++;
                ColNum = quotient;
            } while (!(quotient == 0));
            string alphabetic = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string ans = "";
            for (i = 0; i <= intLetterIndex.Count - 1; i++)
            {
                ans = alphabetic[Convert.ToInt32(intLetterIndex[i] - 1)] + ans;
            }
            return ans;
        }

        /// <summary>
        /// 将Excel表中的字符编号转换为对应的数值
        /// </summary>
        /// <param name="colString">以A1形式表示的列的字母序号，不区分大小写</param>
        /// <returns>以整数的形式返回指定列的数值编号，A列对应数值1</returns>
        /// <remarks></remarks>
        public static int ConvertColumnStringToNumber(string colString)
        {
            colString = colString.ToUpper();
            var ASC_A = Convert.ToInt16('A'); //(byte)(Strings.Asc("A"));
            int ans = 0;
            for (byte i = 0; i <= colString.Length - 1; i++)
            {
                char Chr = colString.ToCharArray(i, 1)[0];
                ans = ans + (Convert.ToInt16(Chr) - ASC_A + 1) * (int)Math.Pow(26, colString.Length - i - 1);
            }
            return ans;
        }

        #endregion

        #region   ---  Range.Value的子数组提取

        /// <summary> 从Range.Value所得到的数组转换为下标值为0的二维数组 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgValue">通过Range.Value提取出来的单元格的数据集合</param>
        /// <returns>与输入的集合相同大小的二维数组，其中第一个元素的下标为0 </returns>
        public static T[,] GetRangeValue<T>(object rgValue)
        {
            object[,] serchingObjects;
            int lowerbound;
            int rowsCount;
            int columnsCount;
            ConvertRangeValueToArray(rgValue, out serchingObjects, out lowerbound, out rowsCount, out columnsCount);

            //  提取所有行与所有列的数据
            int[] rowIndices = new int[rowsCount];
            int[]  columnIndices = new int[columnsCount];
            for (int i = 0; i < rowsCount; i++)
            {
                rowIndices[i] = i;
            }
            for (int i = 0; i < columnsCount; i++)
            {
                columnIndices[i] = i;
            }

            return GetRangeValue2D<T>(serchingObjects, lowerbound, rowIndices, columnIndices);
        }

        /// <summary> 从Range.Value所得到的数组中提取指定的某一行或某一列的数据 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgValue">通过Range.Value提取出来的单元格的数据集合</param>
        /// <param name="isRow">如果是要提取一行值，则为true，否则是提取一列值</param>
        /// <param name="Index">要提取的行或列的下标，第一行的下标为0</param>
        /// <returns>返回一个行向量或者列向量，即返回的二维数组中只有一行值或者一列值。其中第一个元素的下标为0 </returns>
        public static T[,] GetRangeValue<T>(object rgValue, bool isRow, int Index)
        {
            return GetRangeValue<T>(rgValue, isRow, new int[1] { Index });
        }

        /// <summary> 从Range.Value所得到的数组中提取指定的多行或多列的数据 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgValue">通过Range.Value提取出来的单元格的数据集合</param>
        /// <param name="isRow">如果是要提取多个行值，则为true，否则是提取多个列值</param>
        /// <param name="Indices">要提取的行或列的下标，第一行的下标为0</param>
        /// <returns>返回多个行向量或者列向量，其中第一个元素的下标为0 </returns>
        public static T[,] GetRangeValue<T>(object rgValue, bool isRow, int[] Indices)
        {
            object[,] serchingObjects;
            int lowerbound;
            int rowsCount;
            int columnsCount;
            ConvertRangeValueToArray(rgValue, out serchingObjects, out lowerbound, out rowsCount, out columnsCount);
            //
            int[] rowIndices;
            int[] columnIndices;
            //
            if (isRow)  // 提取某一行的数据
            {
                rowIndices = Indices;
                columnIndices = new int[columnsCount];
                for (int i = 0; i < columnsCount; i++)
                {
                    columnIndices[i] = i;
                }
            }
            else  // 提取某一列数据
            {
                columnIndices = Indices;
                rowIndices = new int[rowsCount];
                for (int i = 0; i < rowsCount; i++)
                {
                    rowIndices[i] = i;
                }
            }

            // 提取数据
            return GetRangeValue2D<T>(serchingObjects, lowerbound, rowIndices, columnIndices);
        }

        /// <summary> 从Range.Value所得到的数组中提取指定行与列的新集合 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgValue">通过Range.Value提取出来的单元格的数据集合</param>
        /// <param name="rowIndices">要提取的行的下标，第一行的下标为0</param>
        /// <param name="columnIndices">要提取的列的下标，第一列的下标为0</param>
        /// <returns>提取到的子数组，数组中第一个元素的下标为0。</returns>
        public static T[,] GetRangeValue<T>(object rgValue, int[] rowIndices, int[] columnIndices)
        {
            object[,] serchingObjects;
            int lowerbound;
            int rowsCount;
            int columnsCount;
            ConvertRangeValueToArray(rgValue, out serchingObjects, out lowerbound, out rowsCount, out columnsCount);
            //
            // 提取数据
            return GetRangeValue2D<T>(serchingObjects, lowerbound, rowIndices, columnIndices);
        }

        /// <summary> 从Range.Value所得到的数组（或者单一值）转换为一个二维数组 </summary>
        /// <param name="rgValue"></param>
        /// <param name="convertedArray">转换后的数组，其中第一个元素的下标值并不一定是0，而是lowerBound</param>
        /// <param name="lowerBound">转换后的数组中第一个元素的下标值</param>
        /// <param name="rowsCount">二维数组中的行数</param>
        /// <param name="columnsCount">二维数组中的列数</param>
        private static void ConvertRangeValueToArray(object rgValue, out object[,] convertedArray,
            out int lowerBound, out int rowsCount, out int columnsCount)
        {
            if (rgValue == null) throw new NullReferenceException();

            // 如果只有一个值
            if (!(rgValue is Array))
            {
                convertedArray = new object[1, 1];
                convertedArray[1, 1] = rgValue;
                lowerBound = 0;
                rowsCount = 1;
                columnsCount = 1;
            }
            else
            {
                convertedArray = rgValue as object[,];
                lowerBound = convertedArray.GetLowerBound(0);
                rowsCount = convertedArray.GetUpperBound(0) - lowerBound + 1;
                columnsCount = convertedArray.GetUpperBound(1) - lowerBound + 1;
            }
        }

        /// <summary> 从Range.Value所得到的数组中提取指定行与列的新集合 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serchingObjects">从Range.Value得到的数组，其中第一个元素的下标值并不一定是0，而是lowerBound</param>
        /// <param name="lowerbound"></param>
        /// <param name="rowIndices">要提取的行的下标，第一行的下标为0</param>
        /// <param name="columnIndices">要提取的列的下标，第一列的下标为0</param>
        /// <returns>提取到的子数组，数组中第一个元素的下标为0。</returns>
        private static T[,] GetRangeValue2D<T>(object[,] serchingObjects, int lowerbound, int[] rowIndices, int[] columnIndices)
        {
            // 提取数据
            T[,] ans = new T[rowIndices.Length, columnIndices.Length];
            for (int r = 0; r < rowIndices.Length; r++)
            {
                int rowIndex = rowIndices[r] + lowerbound;
                if (rowIndex > serchingObjects.GetUpperBound(0)) throw new ArgumentOutOfRangeException(@"指定的行号超出集合中行号的最大值");

                for (int c = 0; c < columnIndices.Length; c++)
                {
                    int colIndex = columnIndices[c] + lowerbound;
                    if (colIndex > serchingObjects.GetUpperBound(1)) throw new ArgumentOutOfRangeException(@"指定的列号超出集合中列号的最大值");

                    // 数据的提取与转换
                    ans[r, c] = ConvertExcelValue<T>(serchingObjects[rowIndex, colIndex]);
                }
            }
            return ans;
        }

        /// <summary> 将Excel中的单元格中的值转换为其他类型的值 </summary>
        /// <typeparam name="T">目标类型</typeparam>
        /// <param name="v"></param>
        /// <returns>转换后的值</returns>
        private static T ConvertExcelValue<T>(object v)
        {
            //获取输入的数据类型
            Type destiType = typeof(T);
            TypeCode tpCode = Type.GetTypeCode(destiType);
            //判断此类型的值
            T ans;

            switch (tpCode)
            {
                case TypeCode.DateTime:
                    {
                        try
                        {
                            ans = (T)v;
                        }
                        catch (Exception)
                        {
                            //Debug.Print("数据：" & V.ToString & " 转换为日期出错！将其处理为日期的初始值。"  & vbCrLf & ex.Message)

                            //如果输入的数据为double类型，则将其转换为等效的Date
                            object dt = DateTime.FromOADate(Convert.ToDouble(v));
                            ans = (T)dt;
                        }
                        break;
                    }

                default:
                    {
                        ans = (T)v;
                        break;
                    }
            }
            return ans;
        }
        
        #endregion

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
    }
}