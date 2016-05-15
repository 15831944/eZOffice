Partial Class Ribbon_Wd_zfy
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon_Wd_zfy))
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Btn_TableFormat = Me.Factory.CreateRibbonButton
        Me.CheckBox_DeleteInlineshapes = Me.Factory.CreateRibbonCheckBox
        Me.Gallery1 = Me.Factory.CreateRibbonGallery
        Me.EditBox_Column = Me.Factory.CreateRibbonEditBox
        Me.EditBox_standardString = Me.Factory.CreateRibbonEditBox
        Me.group1 = Me.Factory.CreateRibbonGroup
        Me.Btn_AddBoarder = Me.Factory.CreateRibbonButton
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.btnDeleteRow = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button_SetHyperlinks = Me.Factory.CreateRibbonButton
        Me.Button_ClearTextFormat = Me.Factory.CreateRibbonButton
        Me.btn_ExtractDataFromWordChart = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Button_DeleteSapce = Me.Factory.CreateRibbonButton
        Me.Button_AddSpace = Me.Factory.CreateRibbonButton
        Me.EditBox_SpaceCount = Me.Factory.CreateRibbonEditBox
        Me.Group2.SuspendLayout()
        Me.group1.SuspendLayout()
        Me.Tab1.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Btn_TableFormat)
        Me.Group2.Items.Add(Me.CheckBox_DeleteInlineshapes)
        Me.Group2.Items.Add(Me.Gallery1)
        Me.Group2.Label = "表格"
        Me.Group2.Name = "Group2"
        '
        'Btn_TableFormat
        '
        Me.Btn_TableFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Btn_TableFormat.Label = "表格"
        Me.Btn_TableFormat.Name = "Btn_TableFormat"
        Me.Btn_TableFormat.OfficeImageId = "AdpNewTable"
        Me.Btn_TableFormat.ScreenTip = "规范表格"
        Me.Btn_TableFormat.ShowImage = True
        Me.Btn_TableFormat.SuperTip = "    规范表格，而且删除表格中的嵌入式图片"
        '
        'CheckBox_DeleteInlineshapes
        '
        Me.CheckBox_DeleteInlineshapes.Label = "删除图片"
        Me.CheckBox_DeleteInlineshapes.Name = "CheckBox_DeleteInlineshapes"
        Me.CheckBox_DeleteInlineshapes.ScreenTip = "删除图片"
        Me.CheckBox_DeleteInlineshapes.SuperTip = "    在规范表格时，是否要删除表格中的图片，包括嵌入式或非嵌入式图片。"
        '
        'Gallery1
        '
        Me.Gallery1.Label = "表格样式"
        Me.Gallery1.Name = "Gallery1"
        Me.Gallery1.ScreenTip = "表格样式"
        Me.Gallery1.SuperTip = "    规范表格时所使用的表格样式"
        '
        'EditBox_Column
        '
        Me.EditBox_Column.Label = "列号"
        Me.EditBox_Column.Name = "EditBox_Column"
        Me.EditBox_Column.ScreenTip = "进行检索的字符位于每一行中的第几列。"
        Me.EditBox_Column.Text = "3"
        '
        'EditBox_standardString
        '
        Me.EditBox_standardString.Label = "标志字符"
        Me.EditBox_standardString.Name = "EditBox_standardString"
        Me.EditBox_standardString.ScreenTip = "进行检索判断的字符串"
        Me.EditBox_standardString.Text = "(Inherited from "
        '
        'group1
        '
        Me.group1.Items.Add(Me.Btn_AddBoarder)
        Me.group1.Label = "图形"
        Me.group1.Name = "group1"
        '
        'Btn_AddBoarder
        '
        Me.Btn_AddBoarder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Btn_AddBoarder.Label = "边框"
        Me.Btn_AddBoarder.Name = "Btn_AddBoarder"
        Me.Btn_AddBoarder.OfficeImageId = "AppointmentColor1"
        Me.Btn_AddBoarder.ScreenTip = "嵌入式图片加边框"
        Me.Btn_AddBoarder.ShowImage = True
        Me.Btn_AddBoarder.SuperTip = "    对于非""嵌入式""的图片并没有效果。"
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Label = "eZwd"
        Me.Tab1.Name = "Tab1"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.btnDeleteRow)
        Me.Group4.Items.Add(Me.EditBox_Column)
        Me.Group4.Items.Add(Me.EditBox_standardString)
        Me.Group4.Label = "表格"
        Me.Group4.Name = "Group4"
        '
        'btnDeleteRow
        '
        Me.btnDeleteRow.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnDeleteRow.Label = "删除条目"
        Me.btnDeleteRow.Name = "btnDeleteRow"
        Me.btnDeleteRow.OfficeImageId = "EquationMatrixInsertRowBefore"
        Me.btnDeleteRow.ScreenTip = "删除表格中的特征行"
        Me.btnDeleteRow.ShowImage = True
        Me.btnDeleteRow.SuperTip = " 如果选择的区域中，某一行包含指定的标志字符，则将此行删除。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " 如果选择了一个表格中的多行，则在这些行中进行检索； " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " 如果选择了表格中的某一个单元格，则在这" &
    "一个表格的所有行中进行检索；" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & " 这如果选择了多个表格，则在多个表格中进行检索。"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button_SetHyperlinks)
        Me.Group3.Items.Add(Me.Button_ClearTextFormat)
        Me.Group3.Items.Add(Me.btn_ExtractDataFromWordChart)
        Me.Group3.Label = "文档处理"
        Me.Group3.Name = "Group3"
        '
        'Button_SetHyperlinks
        '
        Me.Button_SetHyperlinks.Label = "网址链接"
        Me.Button_SetHyperlinks.Name = "Button_SetHyperlinks"
        Me.Button_SetHyperlinks.OfficeImageId = "EditHyperlink"
        Me.Button_SetHyperlinks.ScreenTip = "设置网址链接"
        Me.Button_SetHyperlinks.ShowImage = True
        Me.Button_SetHyperlinks.SuperTip = "    此方法的要求是文本的排布格式要求：选择的段落格式必须是：第一段为网页标题，第二段为网址；第三段为网页标题，第四段为网址……，而且其中不能有空行，也不能选择" &
    "空行。"
        '
        'Button_ClearTextFormat
        '
        Me.Button_ClearTextFormat.Label = "清理文本"
        Me.Button_ClearTextFormat.Name = "Button_ClearTextFormat"
        Me.Button_ClearTextFormat.OfficeImageId = "InsertBuildingBlock"
        Me.Button_ClearTextFormat.ScreenTip = "清理文本的格式"
        Me.Button_ClearTextFormat.ShowImage = True
        Me.Button_ClearTextFormat.SuperTip = "    具体过程有： vbcrlf 删除乱码空格、将手动换行符替换为回车、设置嵌入式图片的段落样式"
        '
        'btn_ExtractDataFromWordChart
        '
        Me.btn_ExtractDataFromWordChart.Label = "提取数据"
        Me.btn_ExtractDataFromWordChart.Name = "btn_ExtractDataFromWordChart"
        Me.btn_ExtractDataFromWordChart.OfficeImageId = "ChartTypeXYScatterInsertGallery"
        Me.btn_ExtractDataFromWordChart.ScreenTip = "提取Word中的图表（Chart）数据"
        Me.btn_ExtractDataFromWordChart.ShowImage = True
        Me.btn_ExtractDataFromWordChart.SuperTip = resources.GetString("btn_ExtractDataFromWordChart.SuperTip")
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Button_DeleteSapce)
        Me.Group5.Items.Add(Me.Button_AddSpace)
        Me.Group5.Items.Add(Me.EditBox_SpaceCount)
        Me.Group5.Label = "Coder"
        Me.Group5.Name = "Group5"
        '
        'Button_DeleteSapce
        '
        Me.Button_DeleteSapce.Label = "向前"
        Me.Button_DeleteSapce.Name = "Button_DeleteSapce"
        Me.Button_DeleteSapce.OfficeImageId = "IndentDecrease"
        Me.Button_DeleteSapce.ScreenTip = "代码向左缩进"
        Me.Button_DeleteSapce.ShowImage = True
        Me.Button_DeleteSapce.SuperTip = "删除指定的代码中的前n个空白字符（如果一行中有n个空白字符的话）。"
        '
        'Button_AddSpace
        '
        Me.Button_AddSpace.Label = "向后"
        Me.Button_AddSpace.Name = "Button_AddSpace"
        Me.Button_AddSpace.OfficeImageId = "IndentIncrease"
        Me.Button_AddSpace.ScreenTip = "代码向右缩进"
        Me.Button_AddSpace.ShowImage = True
        Me.Button_AddSpace.SuperTip = "在指定的代码行的开头添加n个空白字符"
        '
        'EditBox_SpaceCount
        '
        Me.EditBox_SpaceCount.Label = "空格数"
        Me.EditBox_SpaceCount.Name = "EditBox_SpaceCount"
        Me.EditBox_SpaceCount.SuperTip = "要在代码行中增加或者删除的空白字符数。"
        Me.EditBox_SpaceCount.Text = "4"
        '
        'Ribbon_Wd_zfy
        '
        Me.Name = "Ribbon_Wd_zfy"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.Tab1)
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.group1.ResumeLayout(False)
        Me.group1.PerformLayout()
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Btn_TableFormat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Btn_AddBoarder As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Gallery1 As Microsoft.Office.Tools.Ribbon.RibbonGallery
    Friend WithEvents CheckBox_DeleteInlineshapes As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button_SetHyperlinks As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button_ClearTextFormat As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_ExtractDataFromWordChart As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeleteRow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox_Column As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents EditBox_standardString As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button_DeleteSapce As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button_AddSpace As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditBox_SpaceCount As Microsoft.Office.Tools.Ribbon.RibbonEditBox

End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon_Wd_zfy
        Get
            Return Me.GetRibbon(Of Ribbon_Wd_zfy)()
        End Get
    End Property
End Class
