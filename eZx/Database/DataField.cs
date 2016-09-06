using System;
using Office = Microsoft.Office.Core;


namespace eZx.Database
{
    /// <summary>
    /// 数据表中每一个字段的信息，包括字段名称，数据类型，在Excel工作表中的列号等
    /// </summary>
    public class DataField
    {
        #region   ---  Properties

        /// <summary> 此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        /// 在Excel 2010中，最大的列号为16384=2^14。 </summary>
        public UInt16 ColumnIndex { get; set; }

        /// <summary> 字段名称 </summary>
        public string Name { get; set; }

        /// <summary> 此列数据的类型 </summary>
        public eZDataType DataType { get; set; }

        /// <summary> 是否允许空值，如果为False，则会自动将其设置为其默认值 </summary>
        public bool NullAllowed { get; set; }

        #endregion

        ///<summary>构造函数</summary>
        /// <param name="fieldName">字段名称</param>
        /// <param name="columnIndex">此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        /// 在Excel 2010中，最大的列号为16384=2^14。 </param>
        /// <param name="dataType">此列的数据类型</param>
        /// <param name="nullAllowed">是否允许有空值</param>
        /// <remarks></remarks>
        public DataField(string fieldName, UInt16 columnIndex, eZDataType dataType = eZDataType.字符, bool nullAllowed = true)
        {
            this.Name = fieldName;
            this.ColumnIndex = columnIndex;
            this.DataType = dataType;
            this.NullAllowed = nullAllowed;
            //
            this.DataType = dataType;
        }

        /// <summary> 检查指定的数据是否符合指定的数据类型 </summary>
        /// <param name="data"> 要进行检测的任意字符 </param>
        /// <param name="ezType"></param>
        /// <returns></returns>
        public static bool IsCompatible(string data, eZDataType ezType)
        {
            bool blnIsCompatible = true;
            switch (ezType)
            {
                case eZDataType.整数:
                    long v_1;
                    blnIsCompatible = long.TryParse(data, out v_1);
                    break;
                case eZDataType.浮点数:
                    double v_2;
                    blnIsCompatible = double.TryParse(data, out v_2);
                    break;
                case eZDataType.日期:
                    DateTime v;
                    blnIsCompatible = DateTime.TryParse(data, out v);
                    break;
            }
            return blnIsCompatible;
        }
    }
}