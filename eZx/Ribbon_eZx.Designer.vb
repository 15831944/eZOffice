Partial Class Ribbon_eZx
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim RibbonDialogLauncherImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher = Me.Factory.CreateRibbonDialogLauncher
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group_DataBase = Me.Factory.CreateRibbonGroup
        Me.btn_DataRange = Me.Factory.CreateRibbonButton
        Me.ButtonValue = Me.Factory.CreateRibbonButton
        Me.btnConstructDatabase = Me.Factory.CreateRibbonButton
        Me.btnEditDatabase = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btn_XYExchange = Me.Factory.CreateRibbonButton
        Me.btn_ExtractDataFromChart = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btnReArrange = Me.Factory.CreateRibbonButton
        Me.EditBox_ReArrangeStart = Me.Factory.CreateRibbonEditBox
        Me.EditBox_ReArrangeEnd = Me.Factory.CreateRibbonEditBox
        Me.EditBox_ReArrangeIntervalId = Me.Factory.CreateRibbonEditBox
        Me.btnShrink = Me.Factory.CreateRibbonButton
        Me.btnReshape = Me.Factory.CreateRibbonButton
        Me.ButtonTranspose = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.EditBox_p1 = Me.Factory.CreateRibbonEditBox
        Me.EditBox_p2 = Me.Factory.CreateRibbonEditBox
        Me.EditBox_p3 = Me.Factory.CreateRibbonEditBox
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.ButtonTest = Me.Factory.CreateRibbonButton
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Group_Help = Me.Factory.CreateRibbonGroup
        Me.btn_ExcelHelp = Me.Factory.CreateRibbonButton
        Me.btn_OfficeHelp = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout
        Me.Group_DataBase.SuspendLayout
        Me.Group1.SuspendLayout
        Me.Group2.SuspendLayout
        Me.Group3.SuspendLayout
        Me.Group4.SuspendLayout
        Me.Tab2.SuspendLayout
        Me.Group_Help.SuspendLayout
        Me.SuspendLayout
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group_DataBase)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "eZx"
        Me.Tab1.Name = "Tab1"
        '
        'Group_DataBase
        '
        Me.Group_DataBase.Items.Add(Me.btn_DataRange)
        Me.Group_DataBase.Items.Add(Me.ButtonValue)
        Me.Group_DataBase.Items.Add(Me.btnConstructDatabase)
        Me.Group_DataBase.Items.Add(Me.btnEditDatabase)
        Me.Group_DataBase.Label = "数据库"
        Me.Group_DataBase.Name = "Group_DataBase"
        '
        'btn_DataRange
        '
        Me.btn_DataRange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_DataRange.Label = "数据范围"
        Me.btn_DataRange.Name = "btn_DataRange"
        Me.btn_DataRange.OfficeImageId = "DatasheetView"
        Me.btn_DataRange.ScreenTip = "选择当前工作表中所有使用到的单元格范围"
        Me.btn_DataRange.ShowImage = true
        '
        'ButtonValue
        '
        Me.ButtonValue.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonValue.Image = Global.ExcelAddIn_zfy.My.Resources.Resources.binary
        Me.ButtonValue.Label = "转换为值"
        Me.ButtonValue.Name = "ButtonValue"
        Me.ButtonValue.ScreenTip = "Range.Value = Range.Value"
        Me.ButtonValue.ShowImage = true
        Me.ButtonValue.SuperTip = "这一操作会将选中的单元格中的公式转化为对应的值，而且，将#DIV/0!、#VALUE!等错误转换为Integer.MinValue"
        '
        'btnConstructDatabase
        '
        Me.btnConstructDatabase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnConstructDatabase.Label = "构造数据库"
        Me.btnConstructDatabase.Name = "btnConstructDatabase"
        Me.btnConstructDatabase.OfficeImageId = "DatabaseSqlServer"
        Me.btnConstructDatabase.ShowImage = true
        '
        'btnEditDatabase
        '
        Me.btnEditDatabase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnEditDatabase.Enabled = false
        Me.btnEditDatabase.Label = "编辑数据库"
        Me.btnEditDatabase.Name = "btnEditDatabase"
        Me.btnEditDatabase.OfficeImageId = "DatabaseSqlServer"
        Me.btnEditDatabase.ShowImage = true
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btn_XYExchange)
        Me.Group1.Items.Add(Me.btn_ExtractDataFromChart)
        Me.Group1.Label = "图表"
        Me.Group1.Name = "Group1"
        '
        'btn_XYExchange
        '
        Me.btn_XYExchange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_XYExchange.Label = "交换XY轴"
        Me.btn_XYExchange.Name = "btn_XYExchange"
        Me.btn_XYExchange.OfficeImageId = "RecoverInviteToMeeting"
        Me.btn_XYExchange.ScreenTip = "交换图表的X轴与Y轴"
        Me.btn_XYExchange.ShowImage = true
        Me.btn_XYExchange.SuperTip = "      对于当前选择的图表，将其中的每一条数据曲线的X数据与Y数据交换，以达到视图上的图表交换XY轴的效果。"
        '
        'btn_ExtractDataFromChart
        '
        Me.btn_ExtractDataFromChart.Label = "提取数据"
        Me.btn_ExtractDataFromChart.Name = "btn_ExtractDataFromChart"
        Me.btn_ExtractDataFromChart.OfficeImageId = "ChartTypeXYScatterInsertGallery"
        Me.btn_ExtractDataFromChart.ScreenTip = "提取图表中的数据"
        Me.btn_ExtractDataFromChart.ShowImage = true
        Me.btn_ExtractDataFromChart.SuperTip = "一般情况下，可以直接通过Excel来提取到Word中的图表中的数据。但是，如果将Excel中的Chart粘贴进Word，而且是以链接的形式粘贴的。在后期操作中，此"& _ 
    "Chart所链接的源Excel文件丢失，此时在Word中便不能直接提取到Excel中的数据了。"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnReArrange)
        Me.Group2.Items.Add(Me.EditBox_ReArrangeStart)
        Me.Group2.Items.Add(Me.EditBox_ReArrangeEnd)
        Me.Group2.Items.Add(Me.EditBox_ReArrangeIntervalId)
        Me.Group2.Items.Add(Me.btnShrink)
        Me.Group2.Items.Add(Me.btnReshape)
        Me.Group2.Items.Add(Me.ButtonTranspose)
        Me.Group2.Label = "数据处理"
        Me.Group2.Name = "Group2"
        '
        'btnReArrange
        '
        Me.btnReArrange.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnReArrange.Label = "数据重排"
        Me.btnReArrange.Name = "btnReArrange"
        Me.btnReArrange.OfficeImageId = "ArrangeTools"
        Me.btnReArrange.ScreenTip = "将选择的数据按指定的区间与间隔进行重新排列"
        Me.btnReArrange.ShowImage = true
        Me.btnReArrange.SuperTip = "用来进行排序的那一列数据只能为数值或者日期"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"如果控制列中的数据不是按递增或者递减的规律排列的，则程序会先将其按大小进行排序。"
        '
        'EditBox_ReArrangeStart
        '
        Me.EditBox_ReArrangeStart.Label = "Start"
        Me.EditBox_ReArrangeStart.Name = "EditBox_ReArrangeStart"
        Me.EditBox_ReArrangeStart.SuperTip = "可以为数值或者日期格式"
        Me.EditBox_ReArrangeStart.Text = Nothing
        '
        'EditBox_ReArrangeEnd
        '
        Me.EditBox_ReArrangeEnd.Label = "End"
        Me.EditBox_ReArrangeEnd.Name = "EditBox_ReArrangeEnd"
        Me.EditBox_ReArrangeEnd.SuperTip = "可以为数值或者日期格式"
        Me.EditBox_ReArrangeEnd.Text = Nothing
        '
        'EditBox_ReArrangeIntervalId
        '
        Me.EditBox_ReArrangeIntervalId.Label = "Interval,Id"
        Me.EditBox_ReArrangeIntervalId.Name = "EditBox_ReArrangeIntervalId"
        Me.EditBox_ReArrangeIntervalId.ScreenTip = "递进步长与用来进行排序的那一列的序号"
        Me.EditBox_ReArrangeIntervalId.SuperTip = "    第一个数值为递进步长，第二个数值为排序数据列，二者用"",""进行分隔。如果是要按选择的单元格区间的第一列来作为进行排序的数据列，则其值为1。"
        Me.EditBox_ReArrangeIntervalId.Text = "1,1"
        '
        'btnShrink
        '
        Me.btnShrink.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnShrink.Label = "消除空行"
        Me.btnShrink.Name = "btnShrink"
        Me.btnShrink.OfficeImageId = "EquationMatrixInsertRowBefore"
        Me.btnShrink.ScreenTip = "将选择的区域中的指定列的元素为空的行的数据删除"
        Me.btnShrink.ShowImage = true
        Me.btnShrink.SuperTip = "注意： 1. 标志列的列号由参数 P1 指定。 "&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"2. 如果单元格有 #VALUE!、#NULL!、#DIV/0!等错误时，会将其处理为Integer类型的最小"& _ 
    "值。"
        '
        'btnReshape
        '
        Me.btnReshape.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnReshape.Label = "表格转换"
        Me.btnReshape.Name = "btnReshape"
        Me.btnReshape.OfficeImageId = "TaskMoveForwardFourWeeks"
        Me.btnReshape.ScreenTip = "将选择的表格重新排列为指定的形式"
        Me.btnReshape.ShowImage = true
        Me.btnReshape.SuperTip = "  类似于Matlab中的 Reshape。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  请在P1中输入新的行数，P2中输入新的列数，在P3中指明是否要将每一列后面的空数据删除（如果数据为空或者为Fa"& _ 
    "lse，则表示不删除结尾空数据）。"&Global.Microsoft.VisualBasic.ChrW(13)&Global.Microsoft.VisualBasic.ChrW(10)&"  在进行重排时，会先将所有的数据的所有列排成一列，然后再一列一列地铺展开来。"
        '
        'ButtonTranspose
        '
        Me.ButtonTranspose.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.ButtonTranspose.Label = "原位转置"
        Me.ButtonTranspose.Name = "ButtonTranspose"
        Me.ButtonTranspose.OfficeImageId = "TableSummarizeWithPivot"
        Me.ButtonTranspose.ScreenTip = "将选中的区域进行原位转置"
        Me.ButtonTranspose.ShowImage = true
        Me.ButtonTranspose.SuperTip = "此命令可以将用户同时选择的多个不相交的小区域分别进行原位转置。"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.EditBox_p1)
        Me.Group3.Items.Add(Me.EditBox_p2)
        Me.Group3.Items.Add(Me.EditBox_p3)
        Me.Group3.Label = "基本参数"
        Me.Group3.Name = "Group3"
        '
        'EditBox_p1
        '
        Me.EditBox_p1.Label = "P1"
        Me.EditBox_p1.Name = "EditBox_p1"
        Me.EditBox_p1.Text = "2"
        '
        'EditBox_p2
        '
        Me.EditBox_p2.Label = "P2"
        Me.EditBox_p2.Name = "EditBox_p2"
        Me.EditBox_p2.ScreenTip = "其他命令的基本参数"
        Me.EditBox_p2.SuperTip = "文本框中的数据类型为Object"
        Me.EditBox_p2.Text = "4"
        '
        'EditBox_p3
        '
        Me.EditBox_p3.Label = "P3"
        Me.EditBox_p3.Name = "EditBox_p3"
        Me.EditBox_p3.Text = "False"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.ButtonTest)
        Me.Group4.Label = "其他"
        Me.Group4.Name = "Group4"
        '
        'ButtonTest
        '
        Me.ButtonTest.Label = "功能测试"
        Me.ButtonTest.Name = "ButtonTest"
        Me.ButtonTest.SuperTip = "在执行此命令之前请自行查看源代码以确认其功能"
        '
        'Tab2
        '
        Me.Tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab2.ControlId.OfficeId = "TabDeveloper"
        Me.Tab2.Groups.Add(Me.Group_Help)
        Me.Tab2.Label = "TabDeveloper"
        Me.Tab2.Name = "Tab2"
        '
        'Group_Help
        '
        Me.Group_Help.DialogLauncher = RibbonDialogLauncherImpl1
        Me.Group_Help.Items.Add(Me.btn_ExcelHelp)
        Me.Group_Help.Items.Add(Me.btn_OfficeHelp)
        Me.Group_Help.Label = "帮助文档"
        Me.Group_Help.Name = "Group_Help"
        Me.Group_Help.Position = Me.Factory.RibbonPosition.AfterOfficeId("GroupXml")
        '
        'btn_ExcelHelp
        '
        Me.btn_ExcelHelp.Label = "Excel开发文档"
        Me.btn_ExcelHelp.Name = "btn_ExcelHelp"
        '
        'btn_OfficeHelp
        '
        Me.btn_OfficeHelp.Label = "Office VBA"
        Me.btn_OfficeHelp.Name = "btn_OfficeHelp"
        '
        'Ribbon_eZx
        '
        Me.Name = "Ribbon_eZx"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(false)
        Me.Tab1.PerformLayout
        Me.Group_DataBase.ResumeLayout(false)
        Me.Group_DataBase.PerformLayout
        Me.Group1.ResumeLayout(false)
        Me.Group1.PerformLayout
        Me.Group2.ResumeLayout(false)
        Me.Group2.PerformLayout
        Me.Group3.ResumeLayout(false)
        Me.Group3.PerformLayout
        Me.Group4.ResumeLayout(false)
        Me.Group4.PerformLayout
        Me.Tab2.ResumeLayout(false)
        Me.Tab2.PerformLayout
        Me.Group_Help.ResumeLayout(false)
        Me.Group_Help.PerformLayout
        Me.ResumeLayout(false)

End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group_DataBase As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_XYExchange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_ExtractDataFromChart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_DataRange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnConstructDatabase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnReArrange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox_ReArrangeStart As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox_ReArrangeEnd As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox_ReArrangeIntervalId As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents btnShrink As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnReshape As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents EditBox_p1 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox_p2 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox_p3 As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnEditDatabase As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group_Help As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_ExcelHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_OfficeHelp As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonTest As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonValue As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonTranspose As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()>
    Friend ReadOnly Property Ribbon_zfy() As Ribbon_eZx
        Get
            Return Me.GetRibbon(Of Ribbon_eZx)()
        End Get
    End Property
End Class
