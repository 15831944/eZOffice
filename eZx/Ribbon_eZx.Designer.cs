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
            this.btn_ToText = this.Factory.CreateRibbonButton();
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
            this.button_SpeedMode = this.Factory.CreateRibbonButton();
            this.button_Interpolations = this.Factory.CreateRibbonButton();
            this.Group3 = this.Factory.CreateRibbonGroup();
            this.EditBox_p1 = this.Factory.CreateRibbonEditBox();
            this.EditBox_p2 = this.Factory.CreateRibbonEditBox();
            this.EditBox_p3 = this.Factory.CreateRibbonEditBox();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button_A3PageSetup = this.Factory.CreateRibbonButton();
            this.button_ContentRowHeight = this.Factory.CreateRibbonButton();
            this.btn_fitToPrint = this.Factory.CreateRibbonButton();
            this.btn_LockSheet = this.Factory.CreateRibbonButton();
            this.btn_UnLockSheet = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.btn_SumupInsertRow = this.Factory.CreateRibbonButton();
            this.btn_DeleteSumRow = this.Factory.CreateRibbonButton();
            this.btn_SepFiles = this.Factory.CreateRibbonButton();
            this.group_slopeProtection = this.Factory.CreateRibbonGroup();
            this.checkBox_ContainsHeader = this.Factory.CreateRibbonCheckBox();
            this.btn_SectionInterp = this.Factory.CreateRibbonButton();
            this.btn_AreaSumup = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn_Station1 = this.Factory.CreateRibbonButton();
            this.btn_Station2 = this.Factory.CreateRibbonButton();
            this.Tab2 = this.Factory.CreateRibbonTab();
            this.Group_Help = this.Factory.CreateRibbonGroup();
            this.btn_ExcelHelp = this.Factory.CreateRibbonButton();
            this.btn_OfficeHelp = this.Factory.CreateRibbonButton();
            this.btn_ExtractCppDoc = this.Factory.CreateRibbonButton();
            this.Tab1.SuspendLayout();
            this.Group_DataBase.SuspendLayout();
            this.Group1.SuspendLayout();
            this.Group2.SuspendLayout();
            this.Group3.SuspendLayout();
            this.group5.SuspendLayout();
            this.group6.SuspendLayout();
            this.group_slopeProtection.SuspendLayout();
            this.group4.SuspendLayout();
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
            this.Tab1.Groups.Add(this.group5);
            this.Tab1.Groups.Add(this.group6);
            this.Tab1.Groups.Add(this.group_slopeProtection);
            this.Tab1.Groups.Add(this.group4);
            this.Tab1.Label = "eZx";
            this.Tab1.Name = "Tab1";
            // 
            // Group_DataBase
            // 
            this.Group_DataBase.Items.Add(this.btn_DataRange);
            this.Group_DataBase.Items.Add(this.ButtonValue);
            this.Group_DataBase.Items.Add(this.btn_ToText);
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
            this.ButtonValue.ScreenTip = "Range.Formula = Range.Formula";
            this.ButtonValue.ShowImage = true;
            this.ButtonValue.SuperTip = "这一操作会将选中的单元格中的公式转化为对应的值，而且，将#DIV/0!、#VALUE!等错误转换为Integer.MinValue";
            this.ButtonValue.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonValue_Click);
            // 
            // btn_ToText
            // 
            this.btn_ToText.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_ToText.Image = ((System.Drawing.Image)(resources.GetObject("btn_ToText.Image")));
            this.btn_ToText.Label = "转换为字符";
            this.btn_ToText.Name = "btn_ToText";
            this.btn_ToText.ScreenTip = "将表格中的值转换为字符";
            this.btn_ToText.ShowImage = true;
            this.btn_ToText.SuperTip = "通过在单元格的值前面添加一个“\'”，来将任意类型的值转换为固定格式的字符。此功能在将表格中的内容粘贴到AutoCAD中的表格时非常有用";
            this.btn_ToText.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_ToText_Click);
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
            this.btn_ExtractDataFromChart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
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
            this.Group2.Items.Add(this.button_SpeedMode);
            this.Group2.Items.Add(this.button_Interpolations);
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
            // button_SpeedMode
            // 
            this.button_SpeedMode.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_SpeedMode.Label = "缩减";
            this.button_SpeedMode.Name = "button_SpeedMode";
            this.button_SpeedMode.OfficeImageId = "ChartTrendline";
            this.button_SpeedMode.ScreenTip = "缩减曲线数据点的数量";
            this.button_SpeedMode.ShowImage = true;
            this.button_SpeedMode.SuperTip = "类似于Origin中的SpeedMode功能，用来将大量数据点曲线缩减为少量的数据点，并保持曲线原来的大致形态。";
            this.button_SpeedMode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_SpeedMode_Click);
            // 
            // button_Interpolations
            // 
            this.button_Interpolations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_Interpolations.Image = global::eZx.Properties.Resources.Interpolation;
            this.button_Interpolations.Label = "样条插值";
            this.button_Interpolations.Name = "button_Interpolations";
            this.button_Interpolations.ScreenTip = "三次样条插值";
            this.button_Interpolations.ShowImage = true;
            this.button_Interpolations.SuperTip = "三次样条的特性为：曲线中任意点处的二阶导数连续。";
            this.button_Interpolations.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_Interpolations_Click);
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
            // 
            // EditBox_p2
            // 
            this.EditBox_p2.Label = "P2";
            this.EditBox_p2.Name = "EditBox_p2";
            this.EditBox_p2.ScreenTip = "其他命令的基本参数";
            this.EditBox_p2.SuperTip = "文本框中的数据类型为Object";
            this.EditBox_p2.Text = "0";
            // 
            // EditBox_p3
            // 
            this.EditBox_p3.Label = "P3";
            this.EditBox_p3.Name = "EditBox_p3";
            this.EditBox_p3.Text = "false";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button_A3PageSetup);
            this.group5.Items.Add(this.button_ContentRowHeight);
            this.group5.Items.Add(this.btn_fitToPrint);
            this.group5.Items.Add(this.btn_LockSheet);
            this.group5.Items.Add(this.btn_UnLockSheet);
            this.group5.Label = "表格规范";
            this.group5.Name = "group5";
            // 
            // button_A3PageSetup
            // 
            this.button_A3PageSetup.Label = "页面设置";
            this.button_A3PageSetup.Name = "button_A3PageSetup";
            this.button_A3PageSetup.ScreenTip = "A3页面打印设置";
            this.button_A3PageSetup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_A3PageSetup_Click);
            // 
            // button_ContentRowHeight
            // 
            this.button_ContentRowHeight.Label = "正文行高";
            this.button_ContentRowHeight.Name = "button_ContentRowHeight";
            this.button_ContentRowHeight.ScreenTip = "A3表格的数据正文的行高设置";
            this.button_ContentRowHeight.SuperTip = "从选定的单元格开始，往下设置一个工作量表的行高";
            this.button_ContentRowHeight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ContentRowHeight_Click);
            // 
            // btn_fitToPrint
            // 
            this.btn_fitToPrint.Label = "对齐打印";
            this.btn_fitToPrint.Name = "btn_fitToPrint";
            this.btn_fitToPrint.ScreenTip = "对齐指定列的边界， 以适应图纸的打印区域";
            this.btn_fitToPrint.SuperTip = "打印右边界的定位值由参数 P1 指定，下边界的定位值由参数 P2 指定，单位为 厘米。\r\n    如果某参数值不大于0，则保持其原定位。\r\n    由于Excel" +
    "中对于列宽的设置是以字符宽度为基准，其值是离散的，所以最终得到的列宽可能与设置的列宽有微小的差别。";
            this.btn_fitToPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_fitToPrint_Click);
            // 
            // btn_LockSheet
            // 
            this.btn_LockSheet.Label = "锁定表格";
            this.btn_LockSheet.Name = "btn_LockSheet";
            this.btn_LockSheet.ScreenTip = "锁定表格";
            this.btn_LockSheet.SuperTip = "如果P3参数为True，则将所有表格锁定";
            this.btn_LockSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_LockSheet_Click);
            // 
            // btn_UnLockSheet
            // 
            this.btn_UnLockSheet.Label = "解锁表格";
            this.btn_UnLockSheet.Name = "btn_UnLockSheet";
            this.btn_UnLockSheet.ScreenTip = "解锁表格";
            this.btn_UnLockSheet.SuperTip = "如果P3参数为True，则将所有表格解锁";
            this.btn_UnLockSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_UnLockSheet_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.btn_SumupInsertRow);
            this.group6.Items.Add(this.btn_DeleteSumRow);
            this.group6.Items.Add(this.btn_SepFiles);
            this.group6.Label = "工程量表";
            this.group6.Name = "group6";
            // 
            // btn_SumupInsertRow
            // 
            this.btn_SumupInsertRow.Label = "插入小计";
            this.btn_SumupInsertRow.Name = "btn_SumupInsertRow";
            this.btn_SumupInsertRow.ScreenTip = "插入小计行";
            this.btn_SumupInsertRow.SuperTip = "对于有很多行数据的工程量表，自动将多数据行进行分隔，并插入小计行";
            this.btn_SumupInsertRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SumupInsertRow_Click);
            // 
            // btn_DeleteSumRow
            // 
            this.btn_DeleteSumRow.Label = "删除小计";
            this.btn_DeleteSumRow.Name = "btn_DeleteSumRow";
            this.btn_DeleteSumRow.ScreenTip = "将同一Sheet中的多个工程量表进行合并";
            this.btn_DeleteSumRow.SuperTip = "删除小计行，并将多个表格中的数据合并";
            this.btn_DeleteSumRow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MergeSumRow_Click);
            // 
            // btn_SepFiles
            // 
            this.btn_SepFiles.Label = "拆分归档";
            this.btn_SepFiles.Name = "btn_SepFiles";
            this.btn_SepFiles.ScreenTip = "将本工作簿中的多个工程量表拆分为单独的工作簿";
            this.btn_SepFiles.SuperTip = "将本工作簿中的多个工程量表拆分为单独的工作簿";
            this.btn_SepFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SepFiles_Click);
            // 
            // group_slopeProtection
            // 
            this.group_slopeProtection.Items.Add(this.checkBox_ContainsHeader);
            this.group_slopeProtection.Items.Add(this.btn_SectionInterp);
            this.group_slopeProtection.Items.Add(this.btn_AreaSumup);
            this.group_slopeProtection.Label = "边坡防护";
            this.group_slopeProtection.Name = "group_slopeProtection";
            this.group_slopeProtection.Visible = false;
            // 
            // checkBox_ContainsHeader
            // 
            this.checkBox_ContainsHeader.Checked = true;
            this.checkBox_ContainsHeader.Label = "包含表头";
            this.checkBox_ContainsHeader.Name = "checkBox_ContainsHeader";
            // 
            // btn_SectionInterp
            // 
            this.btn_SectionInterp.Label = "断面插值";
            this.btn_SectionInterp.Name = "btn_SectionInterp";
            this.btn_SectionInterp.ScreenTip = "对源断面数据进行排序与插值";
            this.btn_SectionInterp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SectionInterp_Click);
            // 
            // btn_AreaSumup
            // 
            this.btn_AreaSumup.Label = "面积汇总";
            this.btn_AreaSumup.Name = "btn_AreaSumup";
            this.btn_AreaSumup.ScreenTip = "对某一种防护方式的面积进行汇总";
            this.btn_AreaSumup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_AreaSumup_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_Station1);
            this.group4.Items.Add(this.btn_Station2);
            this.group4.Items.Add(this.btn_ExtractCppDoc);
            this.group4.Label = "其他操作";
            this.group4.Name = "group4";
            // 
            // btn_Station1
            // 
            this.btn_Station1.Label = "桩号值符";
            this.btn_Station1.Name = "btn_Station1";
            this.btn_Station1.ScreenTip = "将桩号数值转换为字符";
            this.btn_Station1.SuperTip = "转换字符的最大小数位数由参数 P2 指定";
            this.btn_Station1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Station_Click);
            // 
            // btn_Station2
            // 
            this.btn_Station2.Label = "桩号符值";
            this.btn_Station2.Name = "btn_Station2";
            this.btn_Station2.ScreenTip = "将桩号字符转换为数值";
            this.btn_Station2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Station2_Click);
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
            // btn_ExtractCppDoc
            // 
            this.btn_ExtractCppDoc.Label = "C++ 文档提取";
            this.btn_ExtractCppDoc.Name = "btn_ExtractCppDoc";
            this.btn_ExtractCppDoc.ScreenTip = "提取C++文档信息";
            this.btn_ExtractCppDoc.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_ExtractCppDoc_Click);
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
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group_slopeProtection.ResumeLayout(false);
            this.group_slopeProtection.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditDatabase;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab Tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonTranspose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_SpeedMode;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_Interpolations;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Group_Help;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExcelHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_OfficeHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_slopeProtection;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_ContainsHeader;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SectionInterp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_AreaSumup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_fitToPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Station1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_A3PageSetup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_ContentRowHeight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SumupInsertRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_DeleteSumRow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SepFiles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ToText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Station2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_LockSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_UnLockSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_ExtractCppDoc;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_eZx Ribbon_eZx
        {
            get { return this.GetRibbon<Ribbon_eZx>(); }
        }
    }
}
