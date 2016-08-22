namespace eZx
{
    partial class Ribbon_eZx : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_eZx()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon_eZx));
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.Tab1 = this.Factory.CreateRibbonTab();
            this.Group_DataBase = this.Factory.CreateRibbonGroup();
            this.btn_DataRange = this.Factory.CreateRibbonButton();
            this.ButtonValue = this.Factory.CreateRibbonButton();
            this.btnConstructDatabase = this.Factory.CreateRibbonButton();
            this.btnEditDatabase = this.Factory.CreateRibbonButton();
            this.Group1 = this.Factory.CreateRibbonGroup();
            this.btn_XYExchange = this.Factory.CreateRibbonButton();
            this.btn_ExtractDataFromChart = this.Factory.CreateRibbonButton();
            this.Group2 = this.Factory.CreateRibbonGroup();
            this.btnReArrange = this.Factory.CreateRibbonButton();
            this.EditBox_ReArrangeStart = this.Factory.CreateRibbonEditBox();
            this.EditBox_ReArrangeEnd = this.Factory.CreateRibbonEditBox();
            this.EditBox_ReArrangeIntervalId = this.Factory.CreateRibbonEditBox();
            this.btnShrink = this.Factory.CreateRibbonButton();
            this.btnReshape = this.Factory.CreateRibbonButton();
            this.ButtonTranspose = this.Factory.CreateRibbonButton();
            this.Group3 = this.Factory.CreateRibbonGroup();
            this.EditBox_p1 = this.Factory.CreateRibbonEditBox();
            this.EditBox_p2 = this.Factory.CreateRibbonEditBox();
            this.EditBox_p3 = this.Factory.CreateRibbonEditBox();
            this.Group4 = this.Factory.CreateRibbonGroup();
            this.ButtonTest = this.Factory.CreateRibbonButton();
            this.buttonWelcome = this.Factory.CreateRibbonButton();
            this.Tab2 = this.Factory.CreateRibbonTab();
            this.Group_Help = this.Factory.CreateRibbonGroup();
            this.btn_ExcelHelp = this.Factory.CreateRibbonButton();
            this.btn_OfficeHelp = this.Factory.CreateRibbonButton();
            this.Tab1.SuspendLayout();
            this.Group_DataBase.SuspendLayout();
            this.Group1.SuspendLayout();
            this.Group2.SuspendLayout();
            this.Group3.SuspendLayout();
            this.Group4.SuspendLayout();
            this.Tab2.SuspendLayout();
            this.Group_Help.SuspendLayout();
            this.SuspendLayout();
            // 
            // Tab1
            // 
            this.Tab1.Groups.Add(this.Group_DataBase);
            this.Tab1.Groups.Add(this.Group1);
            this.Tab1.Groups.Add(this.Group2);
            this.Tab1.Groups.Add(this.Group3);
            this.Tab1.Groups.Add(this.Group4);
            this.Tab1.Label = "eZx";
            this.Tab1.Name = "Tab1";
            // 
            // Group_DataBase
            // 
            this.Group_DataBase.Items.Add(this.btn_DataRange);
            this.Group_DataBase.Items.Add(this.ButtonValue);
            this.Group_DataBase.Items.Add(this.btnConstructDatabase);
            this.Group_DataBase.Items.Add(this.btnEditDatabase);
            this.Group_DataBase.Label = "数据库";
            this.Group_DataBase.Name = "Group_DataBase";
            // 
            // btn_DataRange
            // 
            this.btn_DataRange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_DataRange.Label = "数据范围";
            this.btn_DataRange.Name = "btn_DataRange";
            this.btn_DataRange.OfficeImageId = "DatasheetView";
            this.btn_DataRange.ScreenTip = "选择当前工作表中所有使用到的单元格范围";
            this.btn_DataRange.ShowImage = true;
            this.btn_DataRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_DataRange_Click);
            // 
            // ButtonValue
            // 
            this.ButtonValue.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonValue.Image = ((System.Drawing.Image)(resources.GetObject("ButtonValue.Image")));
            this.ButtonValue.Label = "转换为值";
            this.ButtonValue.Name = "ButtonValue";
            this.ButtonValue.ScreenTip = "Range.Value = Range.Value";
            this.ButtonValue.ShowImage = true;
            this.ButtonValue.SuperTip = "这一操作会将选中的单元格中的公式转化为对应的值，而且，将#DIV/0!、#VALUE!等错误转换为Integer.MinValue";
            this.ButtonValue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonValue_Click);
            // 
            // btnConstructDatabase
            // 
            this.btnConstructDatabase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConstructDatabase.Label = "构造数据库";
            this.btnConstructDatabase.Name = "btnConstructDatabase";
            this.btnConstructDatabase.OfficeImageId = "DatabaseSqlServer";
            this.btnConstructDatabase.ShowImage = true;
            this.btnConstructDatabase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConstructDatabase_Click);
            // 
            // btnEditDatabase
            // 
            this.btnEditDatabase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEditDatabase.Enabled = false;
            this.btnEditDatabase.Label = "编辑数据库";
            this.btnEditDatabase.Name = "btnEditDatabase";
            this.btnEditDatabase.OfficeImageId = "DatabaseSqlServer";
            this.btnEditDatabase.ShowImage = true;
            this.btnEditDatabase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditDatabase_Click);
            // 
            // Group1
            // 
            this.Group1.Items.Add(this.btn_XYExchange);
            this.Group1.Items.Add(this.btn_ExtractDataFromChart);
            this.Group1.Label = "图表";
            this.Group1.Name = "Group1";
            // 
            // btn_XYExchange
            // 
            this.btn_XYExchange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_XYExchange.Label = "交换XY轴";
            this.btn_XYExchange.Name = "btn_XYExchange";
            this.btn_XYExchange.OfficeImageId = "RecoverInviteToMeeting";
            this.btn_XYExchange.ScreenTip = "交换图表的X轴与Y轴";
            this.btn_XYExchange.ShowImage = true;
            this.btn_XYExchange.SuperTip = "      对于当前选择的图表，将其中的每一条数据曲线的X数据与Y数据交换，以达到视图上的图表交换XY轴的效果。";
            this.btn_XYExchange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_XYExchange_Click);
            // 
            // btn_ExtractDataFromChart
            // 
            this.btn_ExtractDataFromChart.Label = "提取数据";
            this.btn_ExtractDataFromChart.Name = "btn_ExtractDataFromChart";
            this.btn_ExtractDataFromChart.OfficeImageId = "ChartTypeXYScatterInsertGallery";
            this.btn_ExtractDataFromChart.ScreenTip = "提取图表中的数据";
            this.btn_ExtractDataFromChart.ShowImage = true;
            this.btn_ExtractDataFromChart.SuperTip = "一般情况下，可以直接通过Excel来提取到Word中的图表中的数据。但是，如果将Excel中的Chart粘贴进Word，而且是以链接的形式粘贴的。在后期操作中，此" +
    "Chart所链接的源Excel文件丢失，此时在Word中便不能直接提取到Excel中的数据了。";
            this.btn_ExtractDataFromChart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ExtractDataFromChart_Click);
            // 
            // Group2
            // 
            this.Group2.Items.Add(this.btnReArrange);
            this.Group2.Items.Add(this.EditBox_ReArrangeStart);
            this.Group2.Items.Add(this.EditBox_ReArrangeEnd);
            this.Group2.Items.Add(this.EditBox_ReArrangeIntervalId);
            this.Group2.Items.Add(this.btnShrink);
            this.Group2.Items.Add(this.btnReshape);
            this.Group2.Items.Add(this.ButtonTranspose);
            this.Group2.Label = "数据处理";
            this.Group2.Name = "Group2";
            // 
            // btnReArrange
            // 
            this.btnReArrange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReArrange.Label = "数据重排";
            this.btnReArrange.Name = "btnReArrange";
            this.btnReArrange.OfficeImageId = "ArrangeTools";
            this.btnReArrange.ScreenTip = "将选择的数据按指定的区间与间隔进行重新排列";
            this.btnReArrange.ShowImage = true;
            this.btnReArrange.SuperTip = "用来进行排序的那一列数据只能为数值或者日期\r\n如果控制列中的数据不是按递增或者递减的规律排列的，则程序会先将其按大小进行排序。";
            this.btnReArrange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReArrange_Click);
            // 
            // EditBox_ReArrangeStart
            // 
            this.EditBox_ReArrangeStart.Label = "Start";
            this.EditBox_ReArrangeStart.Name = "EditBox_ReArrangeStart";
            this.EditBox_ReArrangeStart.SuperTip = "可以为数值或者日期格式";
            this.EditBox_ReArrangeStart.Text = null;
            // 
            // EditBox_ReArrangeEnd
            // 
            this.EditBox_ReArrangeEnd.Label = "End";
            this.EditBox_ReArrangeEnd.Name = "EditBox_ReArrangeEnd";
            this.EditBox_ReArrangeEnd.SuperTip = "可以为数值或者日期格式";
            this.EditBox_ReArrangeEnd.Text = null;
            // 
            // EditBox_ReArrangeIntervalId
            // 
            this.EditBox_ReArrangeIntervalId.Label = "Interval,Id";
            this.EditBox_ReArrangeIntervalId.Name = "EditBox_ReArrangeIntervalId";
            this.EditBox_ReArrangeIntervalId.ScreenTip = "递进步长与用来进行排序的那一列的序号";
            this.EditBox_ReArrangeIntervalId.SuperTip = "    第一个数值为递进步长，第二个数值为排序数据列，二者用\",\"进行分隔。如果是要按选择的单元格区间的第一列来作为进行排序的数据列，则其值为1。";
            this.EditBox_ReArrangeIntervalId.Text = "1,1";
            // 
            // btnShrink
            // 
            this.btnShrink.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnShrink.Label = "消除空行";
            this.btnShrink.Name = "btnShrink";
            this.btnShrink.OfficeImageId = "EquationMatrixInsertRowBefore";
            this.btnShrink.ScreenTip = "将选择的区域中的指定列的元素为空的行的数据删除";
            this.btnShrink.ShowImage = true;
            this.btnShrink.SuperTip = "注意： 1. 标志列的列号由参数 P1 指定。 \r\n2. 如果单元格有 #VALUE!、#NULL!、#DIV/0!等错误时，会将其处理为Integer类型的最小" +
    "值。";
            this.btnShrink.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShrink_Click);
            // 
            // btnReshape
            // 
            this.btnReshape.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReshape.Label = "表格转换";
            this.btnReshape.Name = "btnReshape";
            this.btnReshape.OfficeImageId = "TaskMoveForwardFourWeeks";
            this.btnReshape.ScreenTip = "将选择的表格重新排列为指定的形式";
            this.btnReshape.ShowImage = true;
            this.btnReshape.SuperTip = "  类似于Matlab中的 Reshape。\r\n  请在P1中输入新的行数，P2中输入新的列数，在P3中指明是否要将每一列后面的空数据删除（如果数据为空或者为Fa" +
    "lse，则表示不删除结尾空数据）。\r\n  在进行重排时，会先将所有的数据的所有列排成一列，然后再一列一列地铺展开来。";
            this.btnReshape.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DataReshape);
            // 
            // ButtonTranspose
            // 
            this.ButtonTranspose.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonTranspose.Label = "原位转置";
            this.ButtonTranspose.Name = "ButtonTranspose";
            this.ButtonTranspose.OfficeImageId = "TableSummarizeWithPivot";
            this.ButtonTranspose.ScreenTip = "将选中的区域进行原位转置";
            this.ButtonTranspose.ShowImage = true;
            this.ButtonTranspose.SuperTip = "此命令可以将用户同时选择的多个不相交的小区域分别进行原位转置。";
            this.ButtonTranspose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTranspose_Click);
            // 
            // Group3
            // 
            this.Group3.Items.Add(this.EditBox_p1);
            this.Group3.Items.Add(this.EditBox_p2);
            this.Group3.Items.Add(this.EditBox_p3);
            this.Group3.Label = "基本参数";
            this.Group3.Name = "Group3";
            // 
            // EditBox_p1
            // 
            this.EditBox_p1.Label = "P1";
            this.EditBox_p1.Name = "EditBox_p1";
            this.EditBox_p1.Text = "2";
            this.EditBox_p1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBox_p1_TextChanged);
            // 
            // EditBox_p2
            // 
            this.EditBox_p2.Label = "P2";
            this.EditBox_p2.Name = "EditBox_p2";
            this.EditBox_p2.ScreenTip = "其他命令的基本参数";
            this.EditBox_p2.SuperTip = "文本框中的数据类型为Object";
            this.EditBox_p2.Text = "4";
            this.EditBox_p2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBox_p2_TextChanged);
            // 
            // EditBox_p3
            // 
            this.EditBox_p3.Label = "P3";
            this.EditBox_p3.Name = "EditBox_p3";
            this.EditBox_p3.Text = "False";
            this.EditBox_p3.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBox_p3_TextChanged);
            // 
            // Group4
            // 
            this.Group4.Items.Add(this.ButtonTest);
            this.Group4.Items.Add(this.buttonWelcome);
            this.Group4.Label = "其他";
            this.Group4.Name = "Group4";
            // 
            // ButtonTest
            // 
            this.ButtonTest.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.ButtonTest.Label = "功能测试";
            this.ButtonTest.Name = "ButtonTest";
            this.ButtonTest.ShowImage = true;
            this.ButtonTest.SuperTip = "在执行此命令之前请自行查看源代码以确认其功能";
            this.ButtonTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonTest_Click);
            // 
            // buttonWelcome
            // 
            this.buttonWelcome.Label = "欢迎";
            this.buttonWelcome.Name = "buttonWelcome";
            this.buttonWelcome.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonWelcome_Click);
            // 
            // Tab2
            // 
            this.Tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Tab2.ControlId.OfficeId = "TabDeveloper";
            this.Tab2.Groups.Add(this.Group_Help);
            this.Tab2.Label = "TabDeveloper";
            this.Tab2.Name = "Tab2";
            // 
            // Group_Help
            // 
            this.Group_Help.DialogLauncher = ribbonDialogLauncherImpl1;
            this.Group_Help.Items.Add(this.btn_ExcelHelp);
            this.Group_Help.Items.Add(this.btn_OfficeHelp);
            this.Group_Help.Label = "帮助文档";
            this.Group_Help.Name = "Group_Help";
            this.Group_Help.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupXml");
            this.Group_Help.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Group_Help_DialogLauncherClick);
            // 
            // btn_ExcelHelp
            // 
            this.btn_ExcelHelp.Label = "Excel开发文档";
            this.btn_ExcelHelp.Name = "btn_ExcelHelp";
            this.btn_ExcelHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ExcelHelp_Click);
            // 
            // btn_OfficeHelp
            // 
            this.btn_OfficeHelp.Label = "Office VBA";
            this.btn_OfficeHelp.Name = "btn_OfficeHelp";
            this.btn_OfficeHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_OfficeHelp_Click);
            // 
            // Ribbon_eZx
            // 
            this.Name = "Ribbon_eZx";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Tab1);
            this.Tabs.Add(this.Tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_zfy_Load);
            this.Tab1.ResumeLayout(false);
            this.Tab1.PerformLayout();
            this.Group_DataBase.ResumeLayout(false);
            this.Group_DataBase.PerformLayout();
            this.Group1.ResumeLayout(false);
            this.Group1.PerformLayout();
            this.Group2.ResumeLayout(false);
            this.Group2.PerformLayout();
            this.Group3.ResumeLayout(false);
            this.Group3.PerformLayout();
            this.Group4.ResumeLayout(false);
            this.Group4.PerformLayout();
            this.Tab2.ResumeLayout(false);
            this.Tab2.PerformLayout();
            this.Group_Help.ResumeLayout(false);
            this.Group_Help.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion


        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_DataBase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_XYExchange;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExtractDataFromChart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_DataRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConstructDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReArrange;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_ReArrangeStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_ReArrangeEnd;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_ReArrangeIntervalId;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShrink;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReshape;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_p1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_p2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox EditBox_p3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Help;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExcelHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_OfficeHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonTranspose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonWelcome;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_eZx Ribbon_eZx
        {
            get { return this.GetRibbon<Ribbon_eZx>(); }
        }
    }
}
