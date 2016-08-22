using System;
using Office = Microsoft.Office.Core;


namespace eZx.Database
{
    /// <summary>
    /// 数据表中每一个字段的信息
    /// </summary>
    public class DataField
    {
        /// <summary> 此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        /// 在Excel 2010中，最大的列号为16384=2^14。 </summary>
        public UInt16 ColumnIndex { get; set; }

        /// <summary> 字段名称 </summary>
        public string Name { get; set; }

        /// <summary> 此列数据的类型 </summary>
        public eZDataType DataType { get; set; }

        /// <summary> 是否允许空值，如果为False，则会自动将其设置为其默认值 </summary>
        public bool NullAllowed { get; set; }

        ///<summary>构造函数</summary>
        /// <param name="name">字段名称</param>
        /// <param name="ColumnIndex">此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        /// 在Excel 2010中，最大的列号为16384=2^14。 </param>
        /// <param name="dataType">此列的数据类型</param>
        /// <param name="nullAllowed">是否允许有空值</param>
        /// <remarks></remarks>
        public DataField(string name, UInt16 ColumnIndex, eZDataType dataType = eZDataType.字符, bool nullAllowed = true)
        {
            DataField with_1 = this;
            with_1.Name = name;
            with_1.ColumnIndex = ColumnIndex;
            with_1.DataType = dataType;
            with_1.NullAllowed = nullAllowed;
            //If fieldtype = Nothing Then
            //    .FieldType = eZDataType.字符
            //End If

            with_1.DataType = dataType;
        }

        public enum eZDataType
        {
            字符, // String
            日期, // DateTime
            整数, // Int64
            浮点数 // Double
        }

        /// <summary> 检查指定的数据是否符合指定的数据类型 </summary>
        /// <param name="CheckedData"></param>
        /// <param name="ezType"></param>
        /// <returns></returns>
        public static bool IsCompatible(string CheckedData, eZDataType ezType)
        {
            bool blnIsCompatible = true;
            switch (ezType)
            {
                case eZDataType.整数:
                    long v_1;
                    blnIsCompatible = long.TryParse(CheckedData, out v_1);
                    break;
                case eZDataType.浮点数:
                    double v_2;
                    blnIsCompatible = double.TryParse(CheckedData, out v_2);
                    break;
                case eZDataType.日期:
                    DateTime v;
                    blnIsCompatible = DateTime.TryParse(CheckedData, out v);
                    break;
            }
            return blnIsCompatible;
        }
    }
}