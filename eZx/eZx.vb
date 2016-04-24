Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.IO
Imports eZstd.eZexcelAPI
Imports eZstd.eZexcelAPI.ExcelExtension

Public Class eZx

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 此Application中所有的数据库工作表
    ''' </summary>
    Private F_DbSheets As New List(Of eZDataSheet)

#End Region

#Region "  ---  Properties"

    Private F_ActiveDataSheet As eZDataSheet
    ''' <summary> 此Application中的活动数据库。 </summary>
    ''' <remarks>如果当前活动的Excel工作表是一个符合格式的数据库工作表，
    ''' 则此属性指向此对应的数据库对象，否则，返回Nothing。</remarks>
    Public Property ActiveDatabaseSheet As eZDataSheet
        Get
            Return Me.F_ActiveDataSheet
        End Get
        Set(value As eZDataSheet)
            F_ActiveDataSheet = value
            If value Is Nothing Then ' 说明此Worksheet不能成功地构成一个数据库格式
                Me.btnEditDatabase.Enabled = False
                Me.btnConstructDatabase.Enabled = True
            Else  ' 说明此Worksheet 符合数据库格式
                Me.btnEditDatabase.Enabled = True
                Me.btnConstructDatabase.Enabled = True
            End If
        End Set
    End Property

#End Region

#Region "  ---  Fields"

    ''' <summary>
    ''' 当前正在运行的Excel程序
    ''' </summary>
    ''' <remarks></remarks>
    Private WithEvents ExcelApp As Excel.Application

    ''' <summary>
    ''' 用来临时保存数据的工作簿
    ''' </summary>
    ''' <remarks>此工作簿用来保存各种临时数据，比如从图表中提取出来的数据情况</remarks>
    Private WithEvents tempWkbk As Excel.Workbook

    ''' <summary>
    ''' 用来临时保存数据的工作簿的文件路径
    ''' </summary>
    ''' <remarks>此工作簿位于桌面上的“tempData.xlsx”</remarks>
    Private path_Tempwkbk = System.IO.Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None), "tempData.xlsx")

    ''' <summary> 供各项命令使用的第一个基本参数 </summary>
    Private Para1 As Object
    ''' <summary> 供各项命令使用的第二个基本参数 </summary>
    Private Para2 As Object
    ''' <summary> 供各项命令使用的第三个基本参数 </summary>
    Private Para3 As Object
#End Region

#End Region

    Private Sub Ribbon_zfy_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ExcelApp = Globals.ThisAddIn.Application
        Para1 = EditBox_p1.Text
        Para2 = EditBox_p2.Text
        Para3 = EditBox_p3.Text
        With Me
            ' .btnEditDatabase.Enabled = False
        End With
    End Sub

#Region "  ---  数据库 ---"

    ''' <summary> 显示工作表中的UsedRange的范围 </summary>
    Private Sub btn_DataRange_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_DataRange.Click
        Dim rg As Excel.Range = ExcelApp.ActiveSheet.UsedRange
        With rg
            .Select()
            ' .Value = .Value '这一操作会将单元格中的公式转化为对应的值，而且，将#DIV/0!、#VALUE!等错误转换为Integer.MinValue
        End With
    End Sub

    ''' <summary>
    ''' 准备构造一个数据库
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub btnConstructDatabase_Click() Handles btnConstructDatabase.Click
        ' -------------------------- 对当前工作表的信息进行处理 --------------------------
        ' 此工作表是否曾经是一个数据库
        Dim sht As Worksheet = ExcelApp.ActiveSheet
        Dim CorrespondingDatasheet As eZDataSheet = CorrespondingInCollection(sht, Me.F_DbSheets)
        Try
            If CorrespondingDatasheet IsNot Nothing Then
                ' 说明此工作表是包含在当前的数据库集合中的，它曾经是一个数据库，但是可能在进行修改后，已经不符合数据库规范了。
                ' ------------ 构造数据库 --------------
                CorrespondingDatasheet = ConstructDatabase()  '将刷新后的数据库更新到集合中的元素中
                Me.ActiveDatabaseSheet = CorrespondingDatasheet
            Else
                ' 说明此工作表并不在数据库集合中，但是它可能是一个数据库。
                ' ------------ 构造数据库 --------------
                Me.ActiveDatabaseSheet = ConstructDatabase()
                If Me.ActiveDatabaseSheet IsNot Nothing Then
                    Me.F_DbSheets.Add(Me.ActiveDatabaseSheet)
                End If
            End If
        Catch ex As NullReferenceException
            MessageBox.Show("当前工作表不符合数据库格式。", "Error", _
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.ActiveDatabaseSheet = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 构造数据库
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ConstructDatabase() As eZDataSheet
        Dim dtSheet As eZDataSheet
        '
        Dim frm As New Form_ConstructDatabase(ExcelApp.ActiveSheet, True)
        dtSheet = frm.ShowDialog()
        '
        Return dtSheet
    End Function

    ''' <summary>
    ''' 准备编辑数据库
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnEditDatabase_Click(sender As Object, e As RibbonControlEventArgs) Handles btnEditDatabase.Click
        Dim dtSheet As eZDataSheet
        '
        Dim frm As New Form_ConstructDatabase(ExcelApp.ActiveSheet, False, Me.ActiveDatabaseSheet)
        dtSheet = frm.ShowDialog()
        '
    End Sub

    ''' <summary>
    ''' 找出某工作表在数据库集合中所对应的那一项，如果没有对应项，则返回Nothing
    ''' </summary>
    ''' <param name="DataSheet">要进行匹配的Excel工作表</param>
    ''' <param name="DatasheetCollection">要进行搜索的数据库集合。</param>
    Private Function CorrespondingInCollection(ByVal DataSheet As Worksheet, _
                     DatasheetCollection As List(Of eZDataSheet)) As eZDataSheet
        Dim dtSheet As eZDataSheet = Nothing
        For Each dbSheet As eZDataSheet In Me.F_DbSheets
            If ExcelFunction.SheetCompare(dbSheet.WorkSheet, DataSheet) Then
                dtSheet = dbSheet
                Exit For
            End If
        Next
        Return dtSheet
    End Function

#End Region

#Region "  ---  图表 ---"

    ''' <summary>
    ''' 交换Excel中活动Chart中的每一条数据曲线的X轴与Y轴
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_XYExchange_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_XYExchange.Click
        Dim ThisChart As Chart = ExcelApp.ActiveChart '当前进行操作的Chart对象
        '
        Static LastChart As Chart ' 上一次进行了“交换XY轴”操作的Chart对象（而不是指上一次激活的Chart对象）
        Static LastX As List(Of Object)
        Static LastY As List(Of Object)
        Static NextExchangeTime As Integer  ' 对于同一个图表，所进行的交换次数，第一次交换时其值为1。
        '
        If ThisChart IsNot Nothing Then
            Dim sr As Series, src As SeriesCollection
            src = ThisChart.SeriesCollection
            If Not ThisChart.Equals(LastChart) Then   '说明是要对一个新的Chart进行操作
                '
                LastX = New List(Of Object)
                LastY = New List(Of Object)
                Dim X As Object, Y As Object
                For Each sr In src
                    X = sr.XValues
                    Y = sr.Values
                    '
                    LastX.Add(X)
                    LastY.Add(Y)
                    '
                    If X.Length > 0 Then
                        sr.XValues = Y
                        sr.Values = X
                    End If
                Next
                NextExchangeTime = 2
            Else  ' 说明还是对原来的那个Chart进行操作
                '此时交换数据时, 使用上一次保存的数据, 而不是直接将现有的Chart中的X与Y交换, 
                '这是因为 : 当X轴为文字，而Y轴为数值时，在交换XY轴后，新的Y轴数据都会变成0，而原来的文字信息在Chart中就不存在了。
                Dim X As Object, Y As Object
                For i = 1 To src.Count
                    sr = src.Item(i)
                    X = LastX.Item(i - 1)
                    Y = LastY.Item(i - 1)
                    If X.Length > 0 Then
                        If NextExchangeTime Mod 2 = 0 Then ' 在偶数次交换时，X与Y列使用其原来的数据
                            sr.XValues = X
                            sr.Values = Y
                        Else
                            sr.XValues = Y
                            sr.Values = X
                        End If
                    End If
                Next
                NextExchangeTime += 1
            End If
            ' 将此次操作的Chart中的数据保存起来
            LastChart = ThisChart
        Else
            MessageBox.Show("没有找到要进行XY轴交换的图表")
        End If
    End Sub

    ''' <summary>
    ''' 提取图表中的数据
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btn_ExtractDataFromChart_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_ExtractDataFromChart.Click
        Dim cht As Chart = ExcelApp.ActiveChart
        '对Chart中的数据进行提取
        If cht IsNot Nothing Then
            ' 打开记录数据的临时工作簿
            If tempWkbk Is Nothing Then
                If File.Exists(path_Tempwkbk) Then
                    tempWkbk = GetObject(path_Tempwkbk)
                Else
                    tempWkbk = ExcelApp.Workbooks.Add()
                    tempWkbk.SaveAs(path_Tempwkbk)
                End If
            End If
            ' 设置写入数据的工作表
            Dim sht As Worksheet = tempWkbk.Worksheets.Item(1)  ' 用工作簿中的第一个工作表来存放数据。
            '
            Dim seriesColl As SeriesCollection = cht.SeriesCollection
            Dim Chartseries As Series
            '开始提取数据
            Dim col As Short = 1
            Dim X As Object, Y As Object, Title As String  ' 这里只能将X与Y的数据类型定义为Object，不能是Object()或者Object(,)
            ' 这里不能用For Each Chartseries in SeriesCollection来引用seriesCollection集合中的元素。
            For i = 1 To seriesColl.Count
                ' 在VB.NET中，seriesCollection集合中的第一个元素的下标值为1。
                Chartseries = seriesColl.Item(i)
                X = Chartseries.XValues
                Y = Chartseries.Values
                Title = Chartseries.Name
                ' 将数据存入Excel表中
                Dim PointsCount As Integer = X.Length
                If PointsCount > 0 Then
                    With sht
                        .Cells(1, col).Value = Title
                        .Range(.Cells(2, col), .Cells(PointsCount + 1, col)).Value = ExcelApp.WorksheetFunction.Transpose(X)
                        .Range(.Cells(2, col + 1), .Cells(PointsCount + 1, col + 1)).Value = ExcelApp.WorksheetFunction.Transpose(Y)
                    End With
                    col = col + 3
                End If
            Next
            tempWkbk.Save()
            With tempWkbk.Application
                .Windows.Item(tempWkbk.Name).Visible = True
                .Windows.Item(tempWkbk.Name).Activate()
                .Visible = True
                If .WindowState = XlWindowState.xlMinimized Then
                    .WindowState = XlWindowState.xlNormal
                End If
            End With

        Else
            MessageBox.Show("没有找到要进行数据提取的图表")
        End If
    End Sub
#End Region

#Region "  ---  数据处理 ---"

    ''' <summary> 进行数据的重新排列 </summary>
    Private Sub btnReArrange_Click(sender As Object, e As RibbonControlEventArgs) Handles btnReArrange.Click


        ' ---------------------------- 确定Range的有效范围 ------------------------------------------
        Dim sht As Worksheet = ExcelApp.ActiveSheet
        Dim rgData As Range = ExcelApp.Selection
        rgData = rgData.Areas.Item(1)
        Dim firstCell As Range  ' 有效区间中的左上角第一个单元
        Dim bottomCell As range ' 有效区间中的左下角的那个单元
        Dim rbcell As Range     ' 有效区间中的右下角的那个单元
        Dim SortedId As Integer, interval As Double
        Dim strInterval_Id() As String = EditBox_ReArrangeIntervalId.Text.Split(",")
        Double.TryParse(strInterval_Id(0), interval)
        Integer.TryParse(strInterval_Id(1), SortedId)
        Dim startRow As Integer
        '
        With rgData
            rbcell = .RBCell
            bottomCell = .Cells(.Rows.Count, SortedId)
            firstCell = .Cells(1, 1)
            If bottomCell.Value Is Nothing Then
                bottomCell = bottomCell.End(XlDirection.xlUp)
            End If
            If firstCell.Value Is Nothing Then
                firstCell = firstCell.End(XlDirection.xlDown)
            End If
            With sht
                rgData = .Range(firstCell, .Cells(bottomCell.Row, rbcell.Column))
            End With
            startRow = rgData.Cells(1, 1).Row
        End With

        ' ------------------------------------- 提取参数 ------------------------------------- 
        Dim rgIdColumn As Range = rgData.Columns(SortedId)
        Dim startData As Double, endData As Double
        Try
            startData = Double.Parse(EditBox_ReArrangeStart.Text)
        Catch ex As Exception
            Try
                startData = Date.Parse(EditBox_ReArrangeStart.Text).ToOADate
            Catch ex1 As Exception
                Try
                    startData = ExcelApp.WorksheetFunction.Min(rgIdColumn)
                    EditBox_ReArrangeStart.Text = startData
                Catch ex2 As Exception
                    MessageBox.Show("指定的数据列中的数据不能进行排序！" & vbCrLf & _
                    ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try
            End Try
        End Try
        Try
            endData = Double.Parse(EditBox_ReArrangeEnd.Text)
        Catch ex As Exception
            Try
                endData = Date.Parse(EditBox_ReArrangeEnd.Text).ToOADate
            Catch ex1 As Exception
                Try
                    endData = ExcelApp.WorksheetFunction.Max(rgIdColumn)
                    EditBox_ReArrangeEnd.Text = endData
                Catch ex2 As Exception
                    MessageBox.Show("指定的数据列中的数据不能进行排序！" & vbCrLf & _
                                ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End Try
            End Try
        End Try

        '
        ' 检查参数的正确性
        If endData <= startData OrElse interval = 0 OrElse interval > (endData - startData) OrElse SortedId = 0 OrElse SortedId > rgData.columns.count Then
            MessageBox.Show("指定的参数不正确！", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' 检查数据的有效性
        Dim Value As Object(,) = rgData.Value2

        Dim v_row As New SortedList(Of Double, Integer)  '每一个key代表此标志列的实际数据，对应的value代表此数据在指定的区间内的局部行号
        Dim r As Integer
        Try
            Dim v As Object
            For r = 1 To UBound(Value, 1)
                v = Value(r, SortedId)
                If (v IsNot Nothing) AndAlso String.Compare("", v.ToString.Trim) <> 0 Then
                    v_row.Add(v, r)
                End If
            Next
        Catch ex As Exception
            Dim c As Range = rgData.Cells(r, SortedId)
            MessageBox.Show("单元格 " & c.Address & " 的数据不符合规范，请检查。" & vbCrLf & _
                            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            c.Activate()
            Exit Sub
        End Try

        ' ------------------------------------------ 开始重新排列数据 ------------------------------------------
        Dim RowsCount As Integer
        Dim ColsCount As Integer = rgData.Columns.Count

        RowsCount = Math.Floor((endData - startData) / interval) + 1

        Dim arrResult(0 To RowsCount - 1, 0 To ColsCount - 1) As Object
        Dim arrKey = v_row.Keys
        Dim arrValue = v_row.Values
        Dim valueRow As Integer
        For r = 0 To RowsCount - 1
            Dim sourceR As Integer = arrKey.IndexOf(startData + r * interval)
            If sourceR >= 0 Then
                valueRow = arrValue.Item(sourceR) ' 指定的数据在Excel区间中的行号
                For c As Integer = 0 To ColsCount - 1
                    arrResult(r, c) = Value(valueRow, c + 1)
                Next
            End If
        Next
        ' 将排列完成后的结果放置回Excel单元格中
        Dim rgResult As Range = sht.Range(firstCell, firstCell.Offset(RowsCount - 1, ColsCount - 1))
        rgResult.Value = arrResult
        rgResult.Columns(SortedId).NumberFormatLocal = rgData.Cells(1, SortedId).NumberFormatLocal  ' 还原这一列的数值格式
        rgResult.Select()
    End Sub

    ''' <summary>
    ''' 消除空行
    ''' </summary>
    Private Sub btnShrink_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShrink.Click

        Dim rgData As Range = ExcelApp.Selection
        rgData = rgData.Areas.Item(1)
        Dim colsCount As Integer = rgData.Columns.Count
        Dim sht As Worksheet = rgData.Worksheet
        '
        Dim SortedId As Integer
        Dim strInterval_Id() As String = EditBox_ReArrangeIntervalId.Text.Split(",")
        If colsCount = 1 Then
            SortedId = 1
        Else
            Integer.TryParse(strInterval_Id(1), SortedId)
            If SortedId = 0 OrElse SortedId > colsCount Then
                MessageBox.Show("指定的数据列的值超出选择的区域范围", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
        '
        Dim firstCell As Range, bottomCell As range, rbcell As Range     ' 有效区间中的右下角的那个单元
        Dim startRow As Integer
        With rgData
            rbcell = .RBCell
            bottomCell = .Cells(.Rows.Count, SortedId)
            firstCell = .Cells(1, 1)
            If bottomCell.Value Is Nothing Then
                bottomCell = bottomCell.End(XlDirection.xlUp)
            End If
            If firstCell.Value Is Nothing Then
                firstCell = firstCell.End(XlDirection.xlDown)
            End If
            With sht
                rgData = .Range(firstCell, .Cells(bottomCell.Row, rbcell.Column))
            End With
            startRow = rgData.Cells(1, 1).Row
        End With
        '

        Dim rowsCount As Integer = rgData.Rows.Count
        Dim arrData(0 To rowsCount - 1, 0 To colsCount - 1) As Object
        '
        Dim Value(,) As Object = rgData.Value

        '
        Dim v As Object
        Dim DataRows As Integer = 0  ' 当前数据行
        For r As Integer = 1 To rowsCount
            v = Value(r, SortedId)
            If (v IsNot Nothing) AndAlso String.Compare("", v.ToString.Trim) <> 0 Then
                For c As Integer = 0 To colsCount - 1
                    arrData(DataRows, c) = Value(r, c + 1)
                Next
                DataRows += 1
            End If
        Next
        ' 将处理完成后的结果放置回Excel单元格中
        Dim rgResult As Range = sht.Range(firstCell, firstCell.Offset(DataRows - 1, colsCount - 1))
        Dim arrResult(0 To DataRows - 1, 0 To colsCount - 1) As Object  ' 剔除无用的数据，而保留非空行
        For r As Integer = 0 To DataRows - 1
            For c As Integer = 0 To colsCount - 1
                arrResult(r, c) = arrData(r, c)
            Next
        Next
        rgResult.Value = arrResult
        rgResult.Select()
    End Sub

    ''' <summary> 数据重排 </summary>
    ''' <remarks>  请在P1中输入新的行数，P2中输入新的列数。 
    ''' 在进行重排时，全先将所有的数据排成一列，然后再进行重排。</remarks>
    Private Sub DataReshape(sender As Object, e As RibbonControlEventArgs) Handles btnReshape.Click
        Dim rg As Excel.Range = ExcelApp.Selection
        Dim startCell As Range = rg.Cells(1, 1)
        Dim Value(,) As Object = rg.Areas.Item(1).Value
        '
        Dim row As UInteger, col As UInteger, blnDeleteNull As Boolean
        Try
            row = Para1
            col = Para2
            blnDeleteNull = (Para3 IsNot Nothing) AndAlso (String.Compare(Para3.ToString, "False", False) <> 0)
            If row = 0 OrElse col = 0 Then
                Throw New ArgumentOutOfRangeException("Col 或 Row", "行或列的数值不能为零。")
            End If
        Catch ex As Exception
            MessageBox.Show("P1或者P2不能转换为数值")
            Exit Sub
        End Try
        '
        Dim ValidDataCount As UInteger  ' 所有数据中，有效的数据的个数
        '将数据由二维表格转换为一维向量，其中只有前面的ValidDataCount个数据是有效的
        Dim arrData() As Object = GetDataListFromTable(Value, blnDeleteNull, ValidDataCount)
        Dim NewShape(0 To row - 1, 0 To col - 1) As Object
        Dim RowIndex, ColIndex As UInteger
        For i As UInteger = 1 To ValidDataCount
            ColIndex = Math.Ceiling(i / row)
            RowIndex = i - (ColIndex - 1) * row
            If i <= row * col Then
                NewShape(RowIndex - 1, ColIndex - 1) = arrData(i - 1)
            Else  ' 考虑到源表格中的有效数据的个数大于目标表格中的元素个数的情况
                Exit For
            End If
        Next
        ' 将重排的数据写入Excel表格中
        Dim DataRg As Range = startCell.Resize(row, col)
        DataRg.Value = NewShape
        DataRg.Select()

    End Sub

    ''' <summary>
    ''' 将Excel中的二维表格数据转换为一个向量
    ''' </summary>
    ''' <param name="Table">要进行数据转换的二维表格</param>
    ''' <param name="DeleteNull">是否要删除每一列结尾处的多个空数据。</param>
    ''' <param name="ValidDataCount">返回的向量中的有效数据的个数，如果DeleteNull的值为False，则其值与二维表格Table中的元素个数相同。</param>
    ''' <returns>一个向量，其中的元素个数与Table中的元素个数相同，但是只有 ValidDataCount 个有效数据</returns>
    ''' <remarks></remarks>
    Private Function GetDataListFromTable(ByRef Table(,) As Object, ByVal DeleteNull As Boolean, _
                                          <System.Runtime.InteropServices.Out> ByRef ValidDataCount As UInteger) As Object()
        Dim Count As Integer = Table.Length
        Dim arrData(0 To Count - 1) As Object
        Dim RowCount As UInteger = UBound(Table, 1) - LBound(Table, 1) + 1
        Dim ColCount As UInteger = UBound(Table, 2) - LBound(Table, 2) + 1
        '
        If DeleteNull Then
            Dim v As Object
            Dim startIndex As UInteger  ' 对于某一列数据而言，其中第一行的数据在转换后的一维向量中的Index
            Dim valueIndex As UInteger  ' 当前要写入的数据在一维向量中的Index
            Dim ValidDataCountInCol As UInteger = 0  ' 本列中有效数据的个数
            For col As UInteger = 1 To ColCount
                For row As UInteger = 1 To RowCount
                    ' 一次处理一列数据
                    v = Table(row, col)
                    valueIndex = startIndex + row - 1
                    arrData(valueIndex) = Table(row, col)  ' 先将这一列的所有数据写入向量中
                    If v IsNot Nothing Then
                        ValidDataCountInCol = row
                    End If
                Next
                startIndex += ValidDataCountInCol
            Next
            ValidDataCount = startIndex  ' 
        Else
            Dim valueIndex As UInteger
            For row As UInteger = 1 To RowCount
                For col As UInteger = 1 To ColCount
                    valueIndex = RowCount * (col - 1) + row
                    arrData(valueIndex - 1) = Table(row, col)
                Next
            Next
            ValidDataCount = RowCount * ColCount
        End If
        '
        Return arrData

    End Function
#End Region

#Region "  ---  其他"

#End Region

#Region "  ---  事件处理 ---"

    ''' <summary>
    '''  激活一个新的工作表
    ''' </summary>
    Private Sub ExcelApp_SheetActivate() Handles ExcelApp.SheetActivate
        Dim sheet As Worksheet = DirectCast(ExcelApp.ActiveSheet, Worksheet)
        Me.ActiveDatabaseSheet = CorrespondingInCollection(sheet, Me.F_DbSheets)
    End Sub

    Private Sub tempWkbk_BeforeClose(ByRef Cancel As Boolean) Handles tempWkbk.BeforeClose
        Me.tempWkbk = Nothing
    End Sub

    Private Sub EditBox_p1_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox_p1.TextChanged
        Dim strText As String = EditBox_p1.Text
        If strText = "" Then
            Para1 = Nothing
        Else
            Para1 = strText
        End If
    End Sub

    Private Sub EditBox_p2_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox_p2.TextChanged
        Dim strText As String = EditBox_p2.Text
        If strText = "" Then
            Para2 = Nothing
        Else
            Para2 = strText
        End If
    End Sub
    Private Sub EditBox_p3_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles EditBox_p3.TextChanged
        Dim strText As String = EditBox_p2.Text
        If strText = "" Then
            Para3 = Nothing
        Else
            Para3 = strText
        End If
    End Sub

#End Region

End Class