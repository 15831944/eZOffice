using System.Collections.Generic;
using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using System.Text;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;

namespace eZwd
{
    partial class Ribbon_eZwd : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_eZwd()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Group2 = this.Factory.CreateRibbonGroup();
            this.Btn_TableFormat = this.Factory.CreateRibbonButton();
            this.CheckBox_DeleteInlineshapes = this.Factory.CreateRibbonCheckBox();
            this.Gallery1 = this.Factory.CreateRibbonGallery();
            this.EditBox_Column = this.Factory.CreateRibbonEditBox();
            this.EditBox_standardString = this.Factory.CreateRibbonEditBox();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Btn_AddBoarder = this.Factory.CreateRibbonButton();
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.Group4 = this.Factory.CreateRibbonGroup();
            this.btnDeleteRow = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btn_ExtractDataFromWordChart = this.Factory.CreateRibbonButton();
            this.Group3 = this.Factory.CreateRibbonGroup();
            this.Button_SetHyperlinks = this.Factory.CreateRibbonButton();
            this.Button_ClearTextFormat = this.Factory.CreateRibbonButton();
            this.buttonPdfReformat = this.Factory.CreateRibbonButton();
            this.btn_CrossRef = this.Factory.CreateRibbonButton();
            this.btn_CrossRefExecute = this.Factory.CreateRibbonButton();
            this.chk_CrossRefReturn = this.Factory.CreateRibbonCheckBox();
            this.Group5 = this.Factory.CreateRibbonGroup();
            this.button_CodeFormater = this.Factory.CreateRibbonButton();
            this.Button_DeleteSapce = this.Factory.CreateRibbonButton();
            this.Button_AddSpace = this.Factory.CreateRibbonButton();
            this.EditBox_SpaceCount = this.Factory.CreateRibbonEditBox();
            this.Group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.Tab1.SuspendLayout();
            this.Group4.SuspendLayout();
            this.group6.SuspendLayout();
            this.Group3.SuspendLayout();
            this.Group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // Group2
            // 
            this.Group2.Items.Add(this.Btn_TableFormat);
            this.Group2.Items.Add(this.CheckBox_DeleteInlineshapes);
            this.Group2.Items.Add(this.Gallery1);
            this.Group2.Label = "表格";
            this.Group2.Name = "Group2";
            // 
            // Btn_TableFormat
            // 
            this.Btn_TableFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_TableFormat.Label = "表格";
            this.Btn_TableFormat.Name = "Btn_TableFormat";
            this.Btn_TableFormat.OfficeImageId = "AdpNewTable";
            this.Btn_TableFormat.ScreenTip = "规范表格";
            this.Btn_TableFormat.ShowImage = true;
            this.Btn_TableFormat.SuperTip = "    规范表格，而且删除表格中的嵌入式图片";
            this.Btn_TableFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_TableFormat_Click);
            // 
            // CheckBox_DeleteInlineshapes
            // 
            this.CheckBox_DeleteInlineshapes.Label = "删除图片";
            this.CheckBox_DeleteInlineshapes.Name = "CheckBox_DeleteInlineshapes";
            this.CheckBox_DeleteInlineshapes.ScreenTip = "删除图片";
            this.CheckBox_DeleteInlineshapes.SuperTip = "    在规范表格时，是否要删除表格中的图片，包括嵌入式或非嵌入式图片。";
            // 
            // Gallery1
            // 
            this.Gallery1.Label = "表格样式";
            this.Gallery1.Name = "Gallery1";
            this.Gallery1.ScreenTip = "表格样式";
            this.Gallery1.SuperTip = "    规范表格时所使用的表格样式";
            this.Gallery1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Gallery1_Click);
            // 
            // EditBox_Column
            // 
            this.EditBox_Column.Label = "列号";
            this.EditBox_Column.Name = "EditBox_Column";
            this.EditBox_Column.ScreenTip = "进行检索的字符位于每一行中的第几列。";
            this.EditBox_Column.Text = "3";
            // 
            // EditBox_standardString
            // 
            this.EditBox_standardString.Label = "标志字符";
            this.EditBox_standardString.Name = "EditBox_standardString";
            this.EditBox_standardString.ScreenTip = "进行检索判断的字符串";
            this.EditBox_standardString.Text = "(Inherited from ";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Btn_AddBoarder);
            this.group1.Label = "图形";
            this.group1.Name = "group1";
            // 
            // Btn_AddBoarder
            // 
            this.Btn_AddBoarder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_AddBoarder.Label = "边框";
            this.Btn_AddBoarder.Name = "Btn_AddBoarder";
            this.Btn_AddBoarder.OfficeImageId = "AppointmentColor1";
            this.Btn_AddBoarder.ScreenTip = "嵌入式图片加边框";
            this.Btn_AddBoarder.ShowImage = true;
            this.Btn_AddBoarder.SuperTip = "    对于非\"嵌入式\"的图片并没有效果。";
            this.Btn_AddBoarder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_AddBoarder_Click);
            // 
            // Tab1
            // 
            this.Tab1.Groups.Add(this.group1);
            this.Tab1.Groups.Add(this.Group2);
            this.Tab1.Groups.Add(this.Group4);
            this.Tab1.Groups.Add(this.group6);
            this.Tab1.Groups.Add(this.Group3);
            this.Tab1.Groups.Add(this.Group5);
            this.Tab1.Label = "eZwd";
            this.Tab1.Name = "Tab1";
            // 
            // Group4
            // 
            this.Group4.Items.Add(this.btnDeleteRow);
            this.Group4.Items.Add(this.EditBox_Column);
            this.Group4.Items.Add(this.EditBox_standardString);
            this.Group4.Label = "表格";
            this.Group4.Name = "Group4";
            // 
            // btnDeleteRow
            // 
            this.btnDeleteRow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDeleteRow.Label = "删除条目";
            this.btnDeleteRow.Name = "btnDeleteRow";
            this.btnDeleteRow.OfficeImageId = "EquationMatrixInsertRowBefore";
            this.btnDeleteRow.ScreenTip = "删除表格中的特征行";
            this.btnDeleteRow.ShowImage = true;
            this.btnDeleteRow.SuperTip = " 如果选择的区域中，某一行包含指定的标志字符，则将此行删除。\r\n 如果选择了一个表格中的多行，则在这些行中进行检索； \r\n 如果选择了表格中的某一个单元格，则在这" +
    "一个表格的所有行中进行检索；\r\n 这如果选择了多个表格，则在多个表格中进行检索。";
            this.btnDeleteRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteRow_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.btn_ExtractDataFromWordChart);
            this.group6.Label = "图表";
            this.group6.Name = "group6";
            // 
            // btn_ExtractDataFromWordChart
            // 
            this.btn_ExtractDataFromWordChart.Label = "提取数据";
            this.btn_ExtractDataFromWordChart.Name = "btn_ExtractDataFromWordChart";
            this.btn_ExtractDataFromWordChart.OfficeImageId = "ChartTypeXYScatterInsertGallery";
            this.btn_ExtractDataFromWordChart.ScreenTip = "提取Word中的图表（Chart）数据";
            this.btn_ExtractDataFromWordChart.ShowImage = true;
            this.btn_ExtractDataFromWordChart.SuperTip = null;
            this.btn_ExtractDataFromWordChart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExtractDataFromWordChart);
            // 
            // Group3
            // 
            this.Group3.Items.Add(this.Button_SetHyperlinks);
            this.Group3.Items.Add(this.Button_ClearTextFormat);
            this.Group3.Items.Add(this.buttonPdfReformat);
            this.Group3.Items.Add(this.btn_CrossRef);
            this.Group3.Items.Add(this.btn_CrossRefExecute);
            this.Group3.Items.Add(this.chk_CrossRefReturn);
            this.Group3.Label = "文档处理";
            this.Group3.Name = "Group3";
            // 
            // Button_SetHyperlinks
            // 
            this.Button_SetHyperlinks.Label = "网址链接";
            this.Button_SetHyperlinks.Name = "Button_SetHyperlinks";
            this.Button_SetHyperlinks.OfficeImageId = "EditHyperlink";
            this.Button_SetHyperlinks.ScreenTip = "设置网址链接";
            this.Button_SetHyperlinks.ShowImage = true;
            this.Button_SetHyperlinks.SuperTip = "    此方法的要求是文本的排布格式要求：选择的段落格式必须是：第一段为网页标题，第二段为网址；第三段为网页标题，第四段为网址……，而且其中不能有空行，也不能选择" +
    "空行。";
            this.Button_SetHyperlinks.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // Button_ClearTextFormat
            // 
            this.Button_ClearTextFormat.Label = "清理文本";
            this.Button_ClearTextFormat.Name = "Button_ClearTextFormat";
            this.Button_ClearTextFormat.OfficeImageId = "InsertBuildingBlock";
            this.Button_ClearTextFormat.ScreenTip = "清理文本的格式";
            this.Button_ClearTextFormat.ShowImage = true;
            this.Button_ClearTextFormat.SuperTip = "    具体过程有： vbcrlf 删除乱码空格、将手动换行符替换为回车、设置嵌入式图片的段落样式";
            this.Button_ClearTextFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_ClearTextFormat_Click);
            // 
            // buttonPdfReformat
            // 
            this.buttonPdfReformat.Label = "段落重排";
            this.buttonPdfReformat.Name = "buttonPdfReformat";
            this.buttonPdfReformat.OfficeImageId = "InsertBuildingBlock";
            this.buttonPdfReformat.ScreenTip = "将多个段落转换为一个段落";
            this.buttonPdfReformat.ShowImage = true;
            this.buttonPdfReformat.SuperTip = "比如将从PDF中粘贴过来的多段文字转换为一个段落。具体操作为：将选择区域的文字中的换行符转换为空格";
            this.buttonPdfReformat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPdfReformat_Click);
            // 
            // btn_CrossRef
            // 
            this.btn_CrossRef.Label = "锚定引用";
            this.btn_CrossRef.Name = "btn_CrossRef";
            this.btn_CrossRef.ScreenTip = "快速引用";
            this.btn_CrossRef.SuperTip = "快速添加交叉引用";
            this.btn_CrossRef.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CrossRef_Click);
            // 
            // btn_CrossRefExecute
            // 
            this.btn_CrossRefExecute.Enabled = false;
            this.btn_CrossRefExecute.Label = "引用";
            this.btn_CrossRefExecute.Name = "btn_CrossRefExecute";
            this.btn_CrossRefExecute.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_CrossRefExecute_Click);
            // 
            // chk_CrossRefReturn
            // 
            this.chk_CrossRefReturn.Enabled = false;
            this.chk_CrossRefReturn.Label = "归位";
            this.chk_CrossRefReturn.Name = "chk_CrossRefReturn";
            // 
            // Group5
            // 
            this.Group5.Items.Add(this.button_CodeFormater);
            this.Group5.Items.Add(this.Button_DeleteSapce);
            this.Group5.Items.Add(this.Button_AddSpace);
            this.Group5.Items.Add(this.EditBox_SpaceCount);
            this.Group5.Label = "Coder";
            this.Group5.Name = "Group5";
            // 
            // button_CodeFormater
            // 
            this.button_CodeFormater.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_CodeFormater.Label = "Format";
            this.button_CodeFormater.Name = "button_CodeFormater";
            this.button_CodeFormater.ScreenTip = "代码格式美化";
            this.button_CodeFormater.ShowImage = true;
            this.button_CodeFormater.SuperTip = "将从Visual Studio 或者 Pycharm中复制到word中的代码进行格式化";
            this.button_CodeFormater.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_CodeFormater_Click);
            // 
            // Button_DeleteSapce
            // 
            this.Button_DeleteSapce.Label = "向前";
            this.Button_DeleteSapce.Name = "Button_DeleteSapce";
            this.Button_DeleteSapce.OfficeImageId = "IndentDecrease";
            this.Button_DeleteSapce.ScreenTip = "代码向左缩进";
            this.Button_DeleteSapce.ShowImage = true;
            this.Button_DeleteSapce.SuperTip = "删除指定的代码中的前n个空白字符（如果一行中有n个空白字符的话）。";
            this.Button_DeleteSapce.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_DeleteSapce_Click);
            // 
            // Button_AddSpace
            // 
            this.Button_AddSpace.Label = "向后";
            this.Button_AddSpace.Name = "Button_AddSpace";
            this.Button_AddSpace.OfficeImageId = "IndentIncrease";
            this.Button_AddSpace.ScreenTip = "代码向右缩进";
            this.Button_AddSpace.ShowImage = true;
            this.Button_AddSpace.SuperTip = "在指定的代码行的开头添加n个空白字符";
            this.Button_AddSpace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button_AddSpace_Click);
            // 
            // EditBox_SpaceCount
            // 
            this.EditBox_SpaceCount.Label = "空格数";
            this.EditBox_SpaceCount.Name = "EditBox_SpaceCount";
            this.EditBox_SpaceCount.SuperTip = "要在代码行中增加或者删除的空白字符数。";
            this.EditBox_SpaceCount.Text = "4";
            // 
            // Ribbon_eZwd
            // 
            this.Name = "Ribbon_eZwd";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.Tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_eZwd_Load);
            this.Group2.ResumeLayout(false);
            this.Group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.Group4.ResumeLayout(false);
            this.Group4.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.Group3.ResumeLayout(false);
            this.Group3.PerformLayout();
            this.Group5.ResumeLayout(false);
            this.Group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_TableFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_AddBoarder;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery Gallery1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox CheckBox_DeleteInlineshapes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_SetHyperlinks;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_ClearTextFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExtractDataFromWordChart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_Column;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_standardString;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_DeleteSapce;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_AddSpace;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_SpaceCount;
        internal RibbonGroup group6;
        internal RibbonButton buttonPdfReformat;
        internal RibbonButton button_CodeFormater;
        internal RibbonButton btn_CrossRef;
        internal RibbonButton btn_CrossRefExecute;
        internal RibbonCheckBox chk_CrossRefReturn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_eZwd Ribbon_eZwd
        {
            get { return this.GetRibbon<Ribbon_eZwd>(); }
        }
    }
}
