Partial Class eZvso
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group_Drawing = Me.Factory.CreateRibbonGroup
        Me.btnPaste = Me.Factory.CreateRibbonButton
        Me.btnArrayCircle = Me.Factory.CreateRibbonButton
        Me.btnArray = Me.Factory.CreateRibbonButton
        Me.btnMove = Me.Factory.CreateRibbonButton
        Me.btnArea = Me.Factory.CreateRibbonButton
        Me.Group_Master = Me.Factory.CreateRibbonGroup
        Me.btnMasterBase = Me.Factory.CreateRibbonSplitButton
        Me.btnLocPin = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group_Drawing.SuspendLayout()
        Me.Group_Master.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group_Drawing)
        Me.Tab1.Groups.Add(Me.Group_Master)
        Me.Tab1.Label = "eZvso"
        Me.Tab1.Name = "Tab1"
        '
        'Group_Drawing
        '
        Me.Group_Drawing.Items.Add(Me.btnPaste)
        Me.Group_Drawing.Items.Add(Me.btnArrayCircle)
        Me.Group_Drawing.Items.Add(Me.btnArray)
        Me.Group_Drawing.Items.Add(Me.btnMove)
        Me.Group_Drawing.Items.Add(Me.btnArea)
        Me.Group_Drawing.Label = "绘图"
        Me.Group_Drawing.Name = "Group_Drawing"
        '
        'btnPaste
        '
        Me.btnPaste.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnPaste.Label = "原位粘贴"
        Me.btnPaste.Name = "btnPaste"
        Me.btnPaste.OfficeImageId = "Paste"
        Me.btnPaste.ShowImage = True
        Me.btnPaste.SuperTip = "    即是通过""开发工具>组合>添加到组""命令实现。" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "    当前选择必须包含要添加的形状和要在其中添加这些形状的组合。组合必须为首要选择或选择中的唯一一个组" &
    "合。"""
        '
        'btnArrayCircle
        '
        Me.btnArrayCircle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnArrayCircle.Label = "阵列"
        Me.btnArrayCircle.Name = "btnArrayCircle"
        Me.btnArrayCircle.OfficeImageId = "PictureBrightnessGallery"
        Me.btnArrayCircle.ShowImage = True
        Me.btnArrayCircle.Tag = "7"
        '
        'btnArray
        '
        Me.btnArray.Label = "阵列"
        Me.btnArray.Name = "btnArray"
        Me.btnArray.OfficeImageId = "NavPaneThumbnailView"
        Me.btnArray.ShowImage = True
        Me.btnArray.Tag = "7"
        '
        'btnMove
        '
        Me.btnMove.Label = "移动"
        Me.btnMove.Name = "btnMove"
        Me.btnMove.OfficeImageId = "PageRightPreview"
        Me.btnMove.ShowImage = True
        Me.btnMove.Tag = "5"
        '
        'btnArea
        '
        Me.btnArea.Label = "面积/周长"
        Me.btnArea.Name = "btnArea"
        Me.btnArea.OfficeImageId = "BlackAndWhiteWhite"
        Me.btnArea.ShowImage = True
        Me.btnArea.Tag = "6"
        '
        'Group_Master
        '
        Me.Group_Master.Items.Add(Me.btnMasterBase)
        Me.Group_Master.Label = "主控形状编辑"
        Me.Group_Master.Name = "Group_Master"
        '
        'btnMasterBase
        '
        Me.btnMasterBase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btnMasterBase.Items.Add(Me.btnLocPin)
        Me.btnMasterBase.Label = "固定基点"
        Me.btnMasterBase.Name = "btnMasterBase"
        Me.btnMasterBase.OfficeImageId = "BorderInside"
        '
        'btnLocPin
        '
        Me.btnLocPin.Label = "局部坐标"
        Me.btnLocPin.Name = "btnLocPin"
        Me.btnLocPin.ShowImage = True
        Me.btnLocPin.SuperTip = "此主控形状的实例对象的旋转中心点相对于实例对象的左下角点的位置"
        '
        'eZvso
        '
        Me.Name = "eZvso"
        Me.RibbonType = "Microsoft.Visio.Drawing"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group_Drawing.ResumeLayout(False)
        Me.Group_Drawing.PerformLayout()
        Me.Group_Master.ResumeLayout(False)
        Me.Group_Master.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group_Master As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnMasterBase As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Group_Drawing As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnPaste As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnMove As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnArray As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnArea As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnArrayCircle As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnLocPin As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property eZvso() As eZvso
        Get
            Return Me.GetRibbon(Of eZvso)()
        End Get
    End Property
End Class
