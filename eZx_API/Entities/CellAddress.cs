using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace eZx_API.Entities
{

    /// <summary>
    /// 指定的单元格在Excel的Worksheet中的位置。左上角第一个单元格的行号值为1，列号值为1。
    /// 可以通过判断其行号或者列号是否为0来判断此类实例是否有赋值。
    /// </summary>
    /// <remarks></remarks>
    public struct CellAddress
    {
        /// <summary>
        /// 单元格在Excel的Worksheet中的行号，左上角第一个单元格的行号值为1。
        /// 在64位的Office 2010中，一个worksheet中共有1048576行，即2^20行。
        /// </summary>
        public UInt32 RowNum;

        /// <summary>
        /// 单元格在Excel的Worksheet中的列号，左上角第一个单元格的列号值为1。
        /// 在64位的Office 2010中，一个worksheet中共有16384列（列号为XFD），即2^14列。
        /// </summary>
        public UInt16 ColNum;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="RowNum_">指定的单元格在Excel的Worksheet中的行号，左上角第一个单元格的行号值为1。</param>
        /// <param name="ColNum_">指定的单元格在Excel的Worksheet中的列号，左上角第一个单元格的列号值为1。</param>
        public CellAddress(UInt32 RowNum_, UInt16 ColNum_)
        {
            RowNum = RowNum_;
            ColNum = ColNum_;
        }
    }
}
