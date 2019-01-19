using System;
using System.Collections.Generic;
using System.Windows.Forms;
using eZstd.Enumerable;
using eZx.AddinManager;
using eZx.Debug;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.CppDoc
{
    /// <summary> 提取C++文档信息，比如MDL开发文档 </summary>
    [EcDescription(CommandDescription)]
    class CppDocExtracter : IExcelExCommand
    {
        #region --- 命令设计

        private const string CommandDescription = @"提取C++文档信息，比如MDL开发文档";

        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new CppDocExtracter();
            return AddinManagerDebuger.DebugInAddinManager(s.ExtractCppDoc,
                excelApp, ref errorMessage, ref errorRange);
        }

        #endregion

        private Application _excelApp;

        /// <summary> </summary>
        public ExternalCommandResult ExtractCppDoc(Application excelApp)
        {
            _excelApp = excelApp;
            excelApp.ScreenUpdating = false;
            var sht = excelApp.ActiveSheet as Worksheet;
            //
            var rg = excelApp.Selection as Range;
            rg = rg.Ex_ShrinkeRange();

            // 设置工作簿的“常规”样式，以确定单元格的行高
            List<CppDocMember> docMembers = ExtractDocMemberFromRange(rg);
            if (docMembers == null || docMembers.Count == 0)
            {
                MessageBox.Show(@"未提取出任何方法，请选择两列多行，且第一个单元格为函数返回值");
                return ExternalCommandResult.Cancelled;
            }
            // 将构造好的数据写入
            Array arr = GetMembersInfo(docMembers);
            Range startCell = rg.Cells[1, 3];
            Range arrRg = RangeValueConverter.FillRange(sht, startCell.Row, startCell.Column, arr);
            // 设计表头行的过滤
            arrRg.Select();
            if (sht.AutoFilterMode)
            {
                // 表示已经打开了过滤
                // arrRg.AutoFilter();
            }
            //
            excelApp.ScreenUpdating = true;
            return ExternalCommandResult.Succeeded;
        }

        /// <summary> 将选择的区域进行解析，得到所有的函数或属性集合 </summary>
        /// <param name="rg"></param>
        /// <returns></returns>
        private List<CppDocMember> ExtractDocMemberFromRange(Range rg)
        {
            var rowsCount = rg.Rows.Count;
            var colsCount = rg.Columns.Count;
            if (colsCount != 2) return null;
            //
            List<CppDocMember> members = new List<CppDocMember>();
            bool extractionStarted = false;
            bool hasDescription = false;
            Range cell = rg[1, 1];
            Range cell_returnType = cell;
            Range cell_Signature = cell;
            Range cell_Description = cell;

            object cellValue;
            // 基本校验
            for (int r = 1; r <= rowsCount; r++)
            {
                cell = rg[r, 1];
                cellValue = cell.Value;
                // 从一个方法的最顶部的特征开始判断

                // 正在进行某一个方法的提取
                if (cellValue != null)
                {
                    if (extractionStarted) // 汇总前一个方法
                    {
                        // 说明已经结束了一个方法的提取
                        AddDocMember(members, cell_returnType, cell_Signature, cell_Description);
                    }
                    // 提取当前的方法
                    cell_returnType = cell;
                    cell_Signature = rg[r, 2];
                    // 初始化所有特征
                    cell_Description = null;
                    hasDescription = false;
                    extractionStarted = true;
                }
                else if (!hasDescription && extractionStarted)
                {
                    // 说明正在提取方法描述
                    cell = rg[r, 2];
                    cell_Description = cell;
                    hasDescription = true;
                }
                if (r == rowsCount)
                {
                    // 说明已经结束了一个方法的提取
                    AddDocMember(members, cell_returnType, cell_Signature, cell_Description);
                }
            }
            return members;
        }

        private void AddDocMember(List<CppDocMember> members, Range cell_returnType, Range cell_Signature, Range cell_Description)
        {
            // 说明已经结束了一个方法的提取
            var memberSignature = (cell_Signature == null) ? null : ((cell_Signature.Value == null ? null : cell_Signature.Value.ToString()));
            var returnType = cell_returnType.Value == null ? null : cell_returnType.Value.ToString();
            var description = (cell_Description == null) ? null : ((cell_Description.Value == null ? null : cell_Description.Value.ToString()));
            if (string.IsNullOrEmpty(memberSignature))
            {
                return;
            }
            var memberName = CppDocMember.ExtractMemberNameFromSignature(memberSignature);
            DocMemberType memberType;
            if (memberName != null)
            {
                memberType = DocMemberType.Function;
            }
            else
            {
                memberType = DocMemberType.Property;
                memberName = memberSignature;
            }
            //
            CppDocMember mem = new CppDocMember(memberName, memberType, memberSignature, returnType,
                description);
            members.Add(mem);
            return;
        }

        /// <summary> 将提取出来的方法对象信息重新构造成数组，以写入到表格中 </summary>
        /// <param name="members"></param>
        /// <returns></returns>
        private Array GetMembersInfo(List<CppDocMember> members)
        {
            var rowsCount = members.Count;
            const int colsCount = 6;
            string[,] memberValue = new string[members.Count + 1, colsCount];
            // 写入表头字段名
            string[] memberTitle = new string[] { "序号", "类型", "返回值", "名称", "描述", "签名", };
            for (int c = 0; c < colsCount; c++)
            {
                memberValue[0, c] = memberTitle[c];
            }
            // 写入数据
            int arrRow;
            for (int r = 0; r < rowsCount; r++)
            {
                arrRow = r + 1;
                var mem = members[r];
                memberValue[arrRow, 0] = arrRow.ToString();
                memberValue[arrRow, 1] = mem.GetMemberTypeName();
                memberValue[arrRow, 2] = mem.ReturnType;
                memberValue[arrRow, 3] = mem.MemberName;
                memberValue[arrRow, 4] = mem.Description;
                memberValue[arrRow, 5] = mem.MemberSignature;
            }
            return memberValue;
        }
    }
}