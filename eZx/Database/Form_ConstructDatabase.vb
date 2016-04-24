Imports System.ComponentModel
Imports Microsoft.Office.Interop.Excel
Imports eZstd.eZexcelAPI
Public Class Form_ConstructDatabase

    Private WithEvents List_FieldInfo As New BindingList(Of DataField)
    Public Property WorkSheet As Worksheet
    ''' <summary>
    ''' 当前窗口是否处于“构造数据库”的模式，如果为False，则为“编辑数据库”的模式
    ''' </summary>
    Private IsConstructingMode As Boolean

    ''' <summary>
    ''' Worksheet.UsedRange.Value所返回的值，此二维数组中，左上角的第一个元素的下标值为(1,1)
    ''' </summary>
    ''' <remarks>此二维数组中包含了字段信息以及每一个字段中的数据</remarks>
    Private F_DataValue(,) As Object

    '''<summary> 此字段名称本身的数据类型。
    ''' 一般情况下，一个字段的名称只要是一个字符就可以了，但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </summary>
    Private F_FieldType As eZDataType

    Dim DataSheet As eZDataSheet = Nothing

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    ''' <summary> 构造函数 </summary>
    ''' <param name="Sheet"></param>
    ''' <param name="ConstructingMode">当前窗口是否处于“构造数据库”的模式，如果为False，则为“编辑数据库”的模式</param>
    '''<param name="DataSheet">当以“构造数据库”模式打开时，此参数可不赋值；当以“编辑数据库”模式打开式，此参数为对应的活动数据库。</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal Sheet As Worksheet, ByVal ConstructingMode As Boolean, Optional DataSheet As eZDataSheet = Nothing)
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        List_FieldInfo.AllowNew = True ' .Add(New DataField("", eZDataType.字符, False, eZDataType.字符))
        '
        Call SetupDataGridView()
        '
        Me.ComboBox_CommonDataType.DataSource = [Enum].GetValues(GetType(eZDataType))
        Me.ComboBox_FieldType.DataSource = [Enum].GetValues(GetType(eZDataType))
        '
        Me.WorkSheet = Sheet
        Me.IsConstructingMode = ConstructingMode
        Me.DataSheet = DataSheet
    End Sub

    Private Sub SetupDataGridView()
        With Me.eZDataGridView1
            .AutoGenerateColumns = False
            .AllowUserToAddRows = True
            .AutoSize = True
            '
            ' 添加数据列
            Dim Column_FieldName As New DataGridViewTextBoxColumn
            With Column_FieldName
                .DataPropertyName = "Name"
                .HeaderText = "名称"
                .Resizable = DataGridViewTriState.False
            End With
            .Columns.Add(Column_FieldName)
            '
            Dim Column_DataType As New DataGridViewComboBoxColumn
            With Column_DataType
                .DataSource = [Enum].GetValues(GetType(eZDataType)) '对于ComboBoxColumn，这一句是必须的。
                .DataPropertyName = "DataType"
                .Name = "DataType"
                .HeaderText = "数据类型"
                .Width = 70
                .Resizable = DataGridViewTriState.False
            End With
            .Columns.Add(Column_DataType)
            '
            Dim Column_NullAllowed As New DataGridViewCheckBoxColumn
            With Column_NullAllowed
                .DataPropertyName = "NullAllowed"
                .Width = 70
                .HeaderText = "允许空值"
                .Resizable = DataGridViewTriState.False
            End With
            .Columns.Add(Column_NullAllowed)
            '
            Dim Column_Check As New DataGridViewButtonColumn
            With Column_Check
                .HeaderText = "检验"
                .Name = "CheckField"
                .Text = "Check Field"
                ' Use the Text property for the button text for all cells rather
                ' than using each cell's value as the text for its own button.
                .UseColumnTextForButtonValue = True
                .Resizable = DataGridViewTriState.False
            End With
            .Columns.Insert(0, Column_Check)
        End With
    End Sub

    Public Overloads Function ShowDialog() As eZDataSheet
        Dim res As DialogResult = MyBase.ShowDialog
        If res = System.Windows.Forms.DialogResult.Yes Then
            ' 构造数据库并返回
            DataSheet = New eZDataSheet(WorkSheet, List_FieldInfo, Me.F_FieldType)

            Return DataSheet
        Else
            Return Nothing
        End If
    End Function

    ' 加载窗口: 每次在Form.ShowDialog方法中，均会触发此Load事件
    Private Sub Form_ConstructDatabase_Load(sender As Object, e As EventArgs) Handles Me.Load
        If IsConstructingMode Then
            Call ConstructDataBase()
        Else
            Call EditDataBase(DataSheet)
        End If
    End Sub

    ' 关闭窗口
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

#End Region

    ''' <summary>
    ''' 构造数据库
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ConstructDataBase()
        Dim rg As Range = WorkSheet.UsedRange
        rg.Select()
        If rg.Cells(1, 1).Address <> WorkSheet.Cells(1, 1).Address Then
            MessageBox.Show("数据表的第一行/列没有数据", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        '
        Me.List_FieldInfo.Clear()
        If rg.Cells.Count > 1 Then
            Me.F_DataValue = rg.Value
            Dim FieldsCount As UShort = UBound(F_DataValue, 2)
            '
            '
            Dim FieldName As String
            For FieldIndex As UShort = 1 To FieldsCount
                FieldName = F_DataValue(1, FieldIndex)
                List_FieldInfo.Add(New DataField(FieldName, FieldIndex))
            Next
        Else ' 说明只选择了一个单元格，此时rg.Value并不会返回一个数组，而是返回一个String或Double等的值
            List_FieldInfo.Add(New DataField(rg.Value.ToString, 1))
        End If
        Me.eZDataGridView1.DataSource = List_FieldInfo
    End Sub

    ''' <summary>
    ''' 编辑数据库
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub EditDataBase(ByVal DataSheet As eZDataSheet)
        Me.Text = "编辑数据库"
        ' 每个字段的数据类型
        Me.eZDataGridView1.DataSource = DataSheet.Fields

        ' 字段名称的数据类型

    End Sub

#Region "  ---  检验字段的信息"
    ''' <summary>
    ''' 同时检验一个字段的名称的数据类型，以及此字段的此列数据的数据类型
    ''' </summary>
    ''' <param name="Field">某一个字段</param>
    ''' <param name="Value">整个数据表的数据（包含字段），数组中的第一个元素的下标为1</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ValidateField(ByVal Field As DataField, ByRef Value As Object(,)) As Boolean
        Dim blnIsValidated As Boolean = True
        If Not ValidateFieldType(Field) Then
            Return False
        End If
        If Not ValidateFieldDataType(Field, Value) Then
            Return False
        End If
        Return blnIsValidated
    End Function

    ''' <summary>
    ''' 检验某一字段的一列数据的数据类型
    ''' </summary>
    ''' <param name="Field">字段信息</param>
    ''' <param name="Value">整个数据表的数据（包含字段），数组中的第一个元素的下标为1</param>
    ''' <returns></returns>
    Private Function ValidateFieldDataType(ByVal Field As DataField, ByRef Value As Object(,)) As Boolean
        Dim blnIsValidated As Boolean = True
        With Field

            Dim DataCount As UInteger = UBound(Value, 1) - 1 ' 数据的个数（不包括字段名称）
            Dim v As Object
            If .NullAllowed Then  ' 允许空值
                For i As UInteger = 2 To DataCount
                    v = Value(i, Field.ColumnIndex)
                    If (v IsNot Nothing) AndAlso (Not IsCompatible(v, .DataType)) Then
                        Return False
                    End If
                Next
            Else  ' 不允许空值
                For i As UInteger = 2 To DataCount
                    If Not IsCompatible(Value(i, Field.ColumnIndex), .DataType) Then
                        Return False
                    End If
                Next

            End If

        End With

        Return blnIsValidated
    End Function

    ''' <summary>
    ''' 检查某一字段的名称本身的数据类型
    ''' </summary>
    ''' <param name="Field"></param>
    ''' <returns></returns>
    Private Function ValidateFieldType(ByVal Field As DataField) As Boolean
        With Field
            If IsCompatible(.Name, F_FieldType) Then
                Return True
            Else
                Return False
            End If
        End With
    End Function
#End Region

#Region "   ---  事件处理"

    ''' <summary> 点击表格控件中的单元格中的对象 </summary>
    Private Sub eZDataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles eZDataGridView1.CellContentClick
        If e.ColumnIndex = 0 Then  ' 说明点击的是“检验字段”的按钮
            Dim FieldDt As DataField = DirectCast(eZDataGridView1.Rows.Item(e.RowIndex).DataBoundItem, DataField)
            If FieldDt.ColumnIndex = 1 Then  ' 第一个字段只检验数据的类型，而不检查字段名称本身的类型
                If ValidateFieldDataType(FieldDt, Me.F_DataValue) Then
                    MessageBox.Show("字段检验合格", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Else
                    MessageBox.Show("字段检验不合格", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                If ValidateField(FieldDt, Me.F_DataValue) Then
                    MessageBox.Show("字段检验合格", "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Else
                    MessageBox.Show("字段检验不合格", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' ! 检验所有的字段，完成数据库的构造或者编辑
    ''' </summary>
    Private Sub CheckAllFields(sender As Object, e As EventArgs) Handles btnCheckAllFields.Click
        Dim FieldDt As DataField
        Dim blnOk As Boolean = True
        If List_FieldInfo.Count > 0 Then
            With Me.eZDataGridView1

                ' 从第二个字段开始来检验字段名称的数据类型，因为对于“字段名称本身的数据类型”的检验，是不包括第一个字段的。
                For Index As Integer = 1 To List_FieldInfo.Count - 1
                    FieldDt = Me.List_FieldInfo.Item(Index)
                    If Not ValidateFieldType(FieldDt) Then
                        MessageBox.Show(String.Format("第{0}个字段： {1} 的字段名称的数据类型检验不合格", Index + 1, FieldDt.Name), _
                                        "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                        blnOk = False
                        Exit For
                    End If
                Next

                ' 从第一个字段开始来检验每一列数据的数据类型
                If blnOk Then
                    For Index As Integer = 0 To List_FieldInfo.Count - 1
                        FieldDt = Me.List_FieldInfo.Item(Index)
                        If Not ValidateFieldDataType(FieldDt, Me.F_DataValue) Then
                            MessageBox.Show(String.Format("第{0}个字段： {1} 的数据检验不合格", Index + 1, FieldDt.Name), _
                                            "Error", MessageBoxButtons.YesNo, MessageBoxIcon.Error)
                            blnOk = False
                            Exit For
                        End If
                    Next

                End If
            End With
            '
        End If
        If blnOk Then
            Me.DialogResult = System.Windows.Forms.DialogResult.Yes
            MessageBox.Show("所有字段检验合格", _
                              "Congratulations", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Me.Close()
        Else
            '   Me.DialogResult = System.Windows.Forms.DialogResult.No
        End If
    End Sub

    ' 错误处理
    Private Sub eZDataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles eZDataGridView1.DataError
        Dim aa = Me.eZDataGridView1.Item(e.ColumnIndex, e.RowIndex).ValueType
        Dim a = Me.eZDataGridView1.Item(e.ColumnIndex, e.RowIndex).Value

        MessageBox.Show(e.Exception.Message & vbCrLf &
                        "行号：" & e.RowIndex & vbCrLf &
                        "列号：" & e.ColumnIndex & vbCrLf &
                        e.Context.ToString)
        e.Cancel = True
    End Sub
    ' 改变基本数据类型
    Private Sub ChangeAllFieldDataType(sender As Object, e As EventArgs) Handles ComboBox_CommonDataType.SelectedValueChanged
        With Me.eZDataGridView1
            Dim ty As eZDataType = DirectCast(Me.ComboBox_CommonDataType.SelectedValue, eZDataType)
            Dim Count As UInteger = .Rows.Count
            For r As Integer = 0 To Count - 1
                .Item("DataType", r).Value = ty
            Next
        End With
    End Sub
    ' 改变字段名称本身的数据类型
    Private Sub ChangeFieldType(sender As Object, e As EventArgs) Handles ComboBox_FieldType.SelectedValueChanged
        With Me.eZDataGridView1
            '  Dim blnSucceed As Boolean
            Dim ezTp As eZDataType = DirectCast(Me.ComboBox_FieldType.SelectedValue, eZDataType)
            Me.F_FieldType = ezTp
            ' 更新界面
            With Me.CheckBox1
                If ezTp = eZDataType.字符 Then
                    .CheckState = CheckState.Indeterminate
                    .Enabled = False
                Else
                    .CheckState = CheckState.Checked
                    .Enabled = True
                End If
            End With

            ' 检验字段
            Dim FieldName As String
            For FieldIndex As UInteger = 1 To List_FieldInfo.Count - 1  ' 不检验第一个字段的数据类型
                Dim df As DataField = List_FieldInfo.Item(FieldIndex)
                FieldName = df.Name
                If Not ValidateFieldType(df) Then
                    MessageBox.Show("第" & df.ColumnIndex & "个字段名称不符合指定的数据类型：" & FieldName, _
                                                    "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
                    '选择出错的那一行
                    Me.eZDataGridView1.Rows(df.ColumnIndex - 1).Selected = True
                    Exit For
                End If
            Next
        End With
    End Sub

    ' 改变每个字段“是否允许空值”
    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim blnAllowNull As Boolean = CheckBox2.Checked
        '
        For Each df As DataField In Me.List_FieldInfo
            df.NullAllowed = blnAllowNull
        Next
        ' 刷新界面显示
        With Me.eZDataGridView1
            .Refresh()
        End With
    End Sub
    ' 添加新的数据行
    Private Sub FieldInfo_AddingNew(sender As Object, e As AddingNewEventArgs) Handles List_FieldInfo.AddingNew
        e.NewObject = New DataField("字段名称", eZDataType.字符, True, eZDataType.字符)
    End Sub

#End Region

End Class