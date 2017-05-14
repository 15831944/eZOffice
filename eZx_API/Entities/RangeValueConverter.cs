using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace eZx_API.Entities
{
    /// <summary> Range.Value的子数组提取 </summary>
    public static class RangeValueConverter
    {
        #region   ---  列号的字符与数值的转换

        /// <summary>
        /// 将Excel表中的列的数值编号转换为对应的字符
        /// </summary>
        /// <param name="colNum">Excel中指定列的数值序号</param>
        /// <returns>以字母序号的形式返回指定列的列号</returns>
        /// <remarks>1对应A；26对应Z；27对应AA</remarks>
        public static string ConvertColumnNumberToString(int colNum)
        {
            // 关键算法就是：连续除以基，直至商为0，从低到高记录余数！
            // 其中value必须是十进制表示的数值
            //intLetterIndex的位数为digits=fix(log(value)/log(26))+1
            //本来算法很简单，但是要解决一个问题：当value为26时，其26进制数为[1 0]，这样的话，
            //以此为下标索引其下面的strAlphabetic时就会出错，因为下标0无法索引。实际上，这种特殊情况下，应该让所得的结果成为[26]，才能索引到字母Z。
            //处理的方法就是，当所得余数remain为零时，就将其改为26，然后将对应的商的值减1.
            if (colNum < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(colNum), @"列数不能小于1");
            }
            List<byte> intLetterIndex = new List<byte>();
            //
            int quotient = 0; //商
            byte remain = 0; //余数
            //
            byte i = (byte)0;
            do
            {
                quotient = (int)Math.Floor((double)colNum / 26); // (int)(Conversion.Fix((double)ColNum / 26)); //商
                remain = (byte)(colNum % 26); //余数
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
                colNum = quotient;
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
            int[] columnIndices = new int[columnsCount];
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
        /// <param name="index">要提取的行或列的下标，第一行的下标为0</param>
        /// <returns>返回一个行向量或者列向量，即返回的二维数组中只有一行值或者一列值。其中第一个元素的下标为0 </returns>
        public static T[,] GetRangeValue<T>(object rgValue, bool isRow, int index)
        {
            return GetRangeValue<T>(rgValue, isRow, new int[1] { index });
        }

        /// <summary> 从Range.Value所得到的数组中提取指定的多行或多列的数据 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rgValue">通过Range.Value提取出来的单元格的数据集合</param>
        /// <param name="isRow">如果是要提取多个行值，则为true，否则是提取多个列值</param>
        /// <param name="indices">要提取的行或列的下标，第一行的下标为0</param>
        /// <returns>返回多个行向量或者列向量，其中第一个元素的下标为0 </returns>
        public static T[,] GetRangeValue<T>(object rgValue, bool isRow, int[] indices)
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
                rowIndices = indices;
                columnIndices = new int[columnsCount];
                for (int i = 0; i < columnsCount; i++)
                {
                    columnIndices[i] = i;
                }
            }
            else  // 提取某一列数据
            {
                columnIndices = indices;
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

        /// <summary> 从Excel中 Range.Value 的大数组中提取子数组 </summary>
        /// <param name="rgValue">Excel中提取出来的单元格数据</param>
        /// <param name="rows">要提取的行号，第一行的下标为0</param>
        /// <param name="cols">要提取的列号，第一列的下标为0</param>
        /// <returns>返回的数组中，第一个元素的下标为0</returns>
        public static object[,] GetRangeValue(object rgValue, int[] rows, int[] cols)
        {
            object[,] result = null;
            if (rgValue is object[,])
            {
                // 则 rgValue 中有不止一个单元格的值
                object[,] rgV = rgValue as object[,];
                int lowBound = rgV.GetLowerBound(0);  // 在 Excel 2016 及以前的版本中，Range.Value提取出来的数组中的第一个元素的下标都为 1。
                // 提取数值
                result = new object[rows.Length, cols.Length];
                for (int r = 0; r < rows.Length; r++)
                {
                    for (int c = 0; c < cols.Length; c++)
                    {
                        result[r, c] = rgV[rows[r + lowBound], cols[r + lowBound]];
                    }
                }
                //
                return result;
            }
            else
            {
                // 则 rgValue 中只有一个单元格的值

                result = new object[1, 1];
                result[1, 1] = rgValue;
                return result;
            }
        }

        /// <summary> 将Excel中 Range.Value的数据转换为二维数组，但不确定数组中第一个元素的下标。 </summary>
        /// <param name="rgValue"> 从Excel中 Range.Value 的大数组 </param>
        /// <returns>如果<paramref name="rgValue"/>中只有一个单元格，则返回的数组的第一个元素的下标必为0；
        /// 如果<paramref name="rgValue"/>中不止一个单元格，则返回的数组的第一个元素的下标很有可能为1；</returns>
        public static object[,] GetRangeValue1(object rgValue)
        {
            if (rgValue is Array)
            {
                // 则 rgValue 有不止一个单元格的值
                return rgValue as object[,]; // 在 Excel 2016 及之前的版本中，此数组 object[,] 中的第一个元素的下标都为 1 。
            }
            else
            {
                // 则 rgValue 中只有一个单元格的值
                object[,] result = new object[1, 1];
                result[0, 0] = rgValue;
                return result;
            }
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

            if (rgValue is Array)
            {
                // 则 rgValue 有不止一个单元格的值
                convertedArray = rgValue as object[,];
                lowerBound = convertedArray.GetLowerBound(0);
                rowsCount = convertedArray.GetUpperBound(0) - lowerBound + 1;
                columnsCount = convertedArray.GetUpperBound(1) - lowerBound + 1;
            }
            else
            {
                // 则 rgValue 中只有一个单元格的值
                convertedArray = new object[1, 1];
                convertedArray[0, 0] = rgValue;
                lowerBound = 0;
                rowsCount = 1;
                columnsCount = 1;

            }
        }

        /// <summary> 从Range.Value所得到的数组中提取指定行与列的新集合 </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="serchingObjects">从Range.Value得到的数组，其中第一个元素的下标值并不一定是0，而是lowerBound</param>
        /// <param name="lowerbound"></param>
        /// <param name="rowIndices">要提取的行的下标，第一行的下标为0</param>
        /// <param name="columnIndices">要提取的列的下标，第一列的下标为0</param>
        /// <returns>提取到的子数组，数组中第一个元素的下标为0。</returns>
        private static T[,] GetRangeValue2D<T>(object[,] serchingObjects, int lowerbound,
            int[] rowIndices, int[] columnIndices)
        {
            // 提取数据
            T[,] ans = new T[rowIndices.Length, columnIndices.Length];
            for (int r = 0; r < rowIndices.Length; r++)
            {
                int rowIndex = rowIndices[r] + lowerbound;
                // if (rowIndex > serchingObjects.GetUpperBound(0)) throw new ArgumentOutOfRangeException(nameof(rowIndex), @"指定的行号超出集合中行号的最大值");

                for (int c = 0; c < columnIndices.Length; c++)
                {
                    int colIndex = columnIndices[c] + lowerbound;
                    // if (colIndex > serchingObjects.GetUpperBound(1)) throw new ArgumentOutOfRangeException(nameof(colIndex), @"指定的列号超出集合中列号的最大值");

                    // 数据的提取与转换
                    ans[r, c] = ConvertExcelValue<T>(serchingObjects[rowIndex, colIndex]);
                }
            }
            return ans;
        }

        /// <summary> 将Excel中的单元格中的值转换为其他类型的值 </summary>
        /// <typeparam name="T">目标类型</typeparam>
        /// <param name="v">Excel中一个单元格中的值</param>
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

        #region   ---  将数组的数据写入到Excel中

        /// <summary> 将数组的数据写入到Excel中 </summary>
        /// <param name="sht"></param>
        /// <param name="startRow">第一行的值为1</param>
        /// <param name="startCol">第一列的值为1</param>
        /// <param name="arr">要写入的数据，为一个一维向量或者二维数组。如果为一维向量，则默认写入一列中。</param>
        /// <param name="colPrior">仅当<paramref name="arr"/>为一维向量，则true表示将数据写入一列，false表示将数据写入一行。
        /// 当<paramref name="arr"/>为二维数组时，此参数没有任何效果</param>
        /// <returns></returns>
        public static void FillRange(Worksheet sht, int startRow, int startCol, Array arr, bool colPrior = true)
        {
            Range rg = GetRange(sht, startRow, startCol, arr, colPrior);
            if (arr.Rank == 1 && colPrior)
            {
                // 如果 arr 为一维向量（比如 int[] arr = new int[4]），则它在Excel中就严格表示一行数据，此时要将其写到一列中，则必须用 Transpose 。
                rg.Value = sht.Application.WorksheetFunction.Transpose(arr);
            }
            else
            {
                rg.Value = arr;
            }
        }

        /// <summary> 根据要写入到Excel中的数组数据确定对应的写入范围 </summary>
        /// <param name="sht"></param>
        /// <param name="startRow">第一行的值为1</param>
        /// <param name="startCol">第一列的值为1</param>
        /// <param name="arr">要写入的数据，为一个一维向量或者二维数组。如果为一维向量，则默认写入一列中。</param>
        /// <param name="colPrior">仅当<paramref name="arr"/>为一维向量，则true表示将数据写入一列，false表示将数据写入一行。
        /// 当<paramref name="arr"/>为二维数组时，此参数没有任何效果</param>
        /// <returns></returns>
        public static Range GetRange(Worksheet sht, int startRow, int startCol, Array arr, bool colPrior = true)
        {
            int rowsCount = arr.GetLength(0);
            int colsCount = 1;
            switch (arr.Rank)
            {
                case 1:  // 表示是一个一维行向量
                    if (colPrior)
                    {
                        colsCount = 1;
                    }
                    else
                    {
                        colsCount = rowsCount;
                        rowsCount = 1;
                    }
                    break;
                case 2:  // 表示是一个二维向量
                    rowsCount = arr.GetLength(0);
                    colsCount = arr.GetLength(1);
                    break;
                default:
                    return null;
            }
            return sht.Range[sht.Cells[startRow, startCol], sht.Cells[startRow + rowsCount - 1, startCol + colsCount - 1]];
        }

        /// <summary> 根据要写入到Excel中的数组数据确定对应的写入范围 </summary>
        /// <param name="sht"></param>
        /// <param name="startRow">第一行的值为1</param>
        /// <param name="startCol">第一列的值为1</param>
        /// <param name="rowsCount">要返回的行数，其值最小为1。</param>
        /// <param name="colsCount">要返回的列数，其值最小为1。</param>
        /// <returns></returns>
        public static Range GetRange(Worksheet sht, int startRow = 1, int startCol = 1, int rowsCount = 1, int colsCount = 1)
        {
            // 构造一个表示区域的字符串，比如 A1:C5
            string rg = $"{ConvertColumnNumberToString(startCol)}{startRow}:{ConvertColumnNumberToString(startCol + colsCount - 1)}{startRow + rowsCount - 1}";
            return sht.Range[rg];
        }

        #endregion
    }
}
