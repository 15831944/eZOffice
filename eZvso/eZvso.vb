Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Visio
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Math
Public Class eZvso

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

#End Region

#Region "  ---  Events"

#End Region

#Region "  ---  Constants"

#End Region

#Region "  ---  Properties"

#End Region

#Region "  ---  Fields"
    Private WithEvents App As Visio.Application
    ''' <summary>
    ''' 当前正在进行编辑的Master对象（不是指Master所对应的实例形状）
    ''' </summary>
    ''' <remarks></remarks>
    Private MasterInEdit As Master

    ''' <summary>
    '''  旋转中心点相对于实例形状范围界定框的左下角点的X位置
    ''' </summary>
    Private MasterBase_LocPinX As Double = 0.5
    ''' <summary>
    ''' 旋转中心点相对于实例形状范围界定框的左下角点的Y位置
    ''' </summary>
    Private MasterBase_LocPinY As Double = 0.5
    ''' <summary>
    ''' 进行放置阵列的对话框
    ''' </summary>
    ''' <remarks></remarks>
    Dim Dlg_CircleArray As Dialog_CircleArray
#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    Private Sub eZvso_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        App = Globals.ThisAddIn.Application
    End Sub

#End Region

#Region "   ---   绘图"

    ''' <summary>
    ''' 将形状原位粘贴到组
    ''' </summary>
    Private Sub AddToGroup(sender As Object, e As RibbonControlEventArgs) Handles btnPaste.Click
        App.ActiveWindow.Selection.AddToGroup()
    End Sub

    ''' <summary>
    ''' 图形的平移、矩形阵列、面积与周长
    ''' </summary>
    Private Sub btnMove_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMove.Click, btnArray.Click, btnArea.Click
        Dim btn As RibbonButton = DirectCast(sender, RibbonButton)
        App.Addons.Item(Short.Parse(btn.Tag)).Run("")  '5 表示平移，6表示形状的面积与周长，7表示阵列
        '如果要执行阵列命令，也可以用：    App.DoCmd(1354)  ' VisUICmds 常量中的 visCmdToolsArrayShapesAddOn  命令
    End Sub

    ''' <summary>
    ''' 旋转阵列
    ''' </summary>
    Private Sub CircleArray(sender As Object, e As RibbonControlEventArgs) Handles btnArrayCircle.Click
        Dim sel As Selection = App.ActiveWindow.Selection
        If sel.Count > 0 Then
            If Dlg_CircleArray Is Nothing Then
                Dlg_CircleArray = New Dialog_CircleArray
            End If
            Dim angle As Double     ' 旋转阵列的总角度
            Dim n As UShort     ' 旋转阵列的个数
            Dim blnPreserveDirection As Boolean  '是否保留对象的角度方向
            '
            Dim res As DialogResult = Dlg_CircleArray.ShowDialog(Num:=n, Angle:=angle, blnPreserveDirection:=blnPreserveDirection)
            If res = DialogResult.OK Then
                Dim shp As Visio.Shape
                If sel.Count = 1 Then
                    shp = sel.Item(1)
                Else
                    shp = sel.Group
                End If
                With shp
                    Dim baseX As Double = .Cells("PinX").ResultIU   ' 图形的旋转中心在页面中的绝对X坐标
                    Dim baseY As Double = .Cells("PinY").ResultIU  ' 图形的旋转中心在页面中的绝对Y坐标
                    Dim OrigionalAngle As Double = .Cells("Angle").Result(Visio.VisUnitCodes.visDegrees)
                    Dim Width As Double = .Cells("Width").ResultIU
                    Dim Height As Double = .Cells("Height").ResultIU
                    Dim strLocPinX As String = .Cells("LocPinX").Formula
                    Dim strLocPinY As String = .Cells("LocPinY").Formula
                    Dim WidthScale As Double, HeightScale As Double

                    Try
                        WidthScale = Double.Parse(strLocPinX.Substring(6))  ' 图形的旋转中心相对于图形的左下角点的位置： Width*0.5
                        HeightScale = Double.Parse(strLocPinY.Substring(7)) ' 图形的旋转中心相对于图形的左下角点的位置： Height*0.5
                    Catch ex As Exception
                        MessageBox.Show("请在ShapeSheet中以相对值的形式来表达 LocPinX 与 LocPinY 的值", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                    Dim OriginalCenterX As Double = baseX - Width * WidthScale + Width * 0.5 ' 图形的中心点在页面中的绝对X坐标
                    Dim OriginalCenterY As Double = baseY - Height * HeightScale + Height * 0.5 ' 图形的中心点在页面中的绝对Y坐标
                    Dim r As Double = Sqrt((baseX - OriginalCenterX) ^ 2 + (baseY - OriginalCenterY) ^ 2)
                    ' ------------------------ 开始复制形状  ----------------------
                    Dim NewShape As Shape
                    App.ShowChanges = False
                    For i As UShort = 1 To n - 1
                        Dim deltaA As Double = (angle / n * i) / 180 * PI  ' 单位为弧度
                        NewShape = shp.Duplicate()
                        With NewShape
                            '将形状移动回原位
                            .Cells("PinX").ResultIU = baseX
                            .Cells("PinY").ResultIU = baseY
                            If blnPreserveDirection Then   '是否保留对象的角度方向
                                If r > 0 Then   ' 此时新图形与原图形在同一个位置，不用作任何的移动，而且下面的alpha角算出来为无穷，因为分母r为0.
                                    Dim alpha As Double = Asin((OriginalCenterY - baseY) / r)
                                    Dim NewCenterX As Double = baseX + r * Cos(deltaA + alpha) ' 注意三角函数计算时的单位为弧度
                                    Dim NewCenterY As Double = baseY + r * Sin(deltaA + alpha)
                                    '新形状的中心点在页面中的绝对坐标值
                                    .Cells("PinX").ResultIU = NewCenterX - OriginalCenterX + baseX
                                    .Cells("PinY").ResultIU = NewCenterY - OriginalCenterY + baseY
                                End If
                            Else
                                .Cells("Angle").Result(Visio.VisUnitCodes.visDegrees) = OrigionalAngle + deltaA / PI * 180
                            End If
                        End With
                    Next


                    App.ShowChanges = True
                End With
            End If
        End If

    End Sub

#End Region

#Region "   ---   主控形状编辑"

    ''' <summary>
    ''' 绘制主控形状的基点位置
    ''' </summary>
    Private Sub btnMasterBase_Click(sender As Object, e As RibbonControlEventArgs) Handles btnMasterBase.Click
        With MasterInEdit
            Dim shape As Shape = .DrawLine(0, 0, 0, 0)
            ' MasterInEdit.

        End With
    End Sub

    Private Sub btnLocPin_Click(sender As Object, e As RibbonControlEventArgs) Handles btnLocPin.Click

    End Sub


#End Region

#Region "   ---   事件处理"

    Private Sub App_WindowActivated(Window As Window) Handles App.WindowActivated
        If Window.SubType = 64 Then  ' 主控形状绘图页窗口。（通过“文档模具”中右键，“编辑主控形状”所进入的窗口）
            Me.btnMasterBase.Enabled = True
            MasterInEdit = DirectCast(Window.Master, Master)
        Else
            Me.btnMasterBase.Enabled = False
            MasterInEdit = Nothing
        End If
    End Sub
#End Region

End Class
