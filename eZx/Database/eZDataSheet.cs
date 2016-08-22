using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace eZx.Database
{
    public class eZDataSheet
    {
        #region   ---  Declarations & Definitions


        #region   ---  Properties

        /// <summary>
        /// 代表此数据库的工作表对象
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public Worksheet WorkSheet { get; set; }

        /// <summary>
        /// 此工作表是否符合数据库的格式规范
        /// </summary>
        private bool F_IsFormated;

        public BindingList<DataField> Fields { get; set; }

        ///<summary> 此字段名称本身的数据类型。
        /// 一般情况下，一个字段的名称只要是一个字符就可以了，但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </summary>
        public DataField.eZDataType FieldType { get; set; }

        #endregion

        #region   ---  Fields

        #endregion

        #endregion

        #region   ---  构造函数与窗体的加载、打开与关闭

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="WkSheet"></param>
        /// <param name="List_FieldInfo"></param>
        /// <param name="FieldType">字段名称本身的数据类型, 一般情况下，一个字段的名称只要是一个字符就可以了，
        /// 但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </param>
        /// <remarks></remarks>
        public eZDataSheet(Worksheet WkSheet, BindingList<DataField> List_FieldInfo,
            DataField.eZDataType FieldType = DataField.eZDataType.字符)
        {
            this.WorkSheet = WkSheet;
            this.Fields = List_FieldInfo;
            this.FieldType = FieldType;
        }

        #endregion
    }
}