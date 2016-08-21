Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop
Imports System.IO
Public Class Ribbon_Wd_zfy

#Region "  ---  Declarations & Definitions"

#Region "  ---  Events"

#End Region

#Region "  ---  Properties"

#End Region

#Region "  ---  Fields"
    Private App As Application
    ''' <summary>
    ''' 当前正在运行的Word程序中的活动Word文档
    ''' </summary>
    ''' <remarks></remarks>
    Private Doc As Document

    ''' <summary>
    ''' 进行表格规范化时所使用的表格样式
    ''' </summary>
    ''' <remarks>
    ''' 注意：在为内容指定样式（比如为段落指定段落样式或者为表格指定表格样式）时，
    ''' 如果指定的样式不存在或者为段落指定了表格样式等时，程序会继续正常执行，也不会跳过后面的语句，
    ''' 只是就相当于没有执行这一行。</remarks>
    Private F_TableStyle As String = "zengfy表格-上下总分型1"

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        App = Globals.ThisAddIn.Application
        Doc = App.ActiveDocument
        Call ListStyles(Doc, Me.Gallery1)
    End Sub

#End Region

#Region "  ---  界面操作"
    '列出与选择表格样式
    ''' <summary>
    ''' 列出文档中所有的表格样式
    ''' </summary>
    ''' <param name="doc"></param>
    ''' <param name="Gallary"></param>
    ''' <remarks></remarks>
    Private Sub ListStyles(ByVal doc As Document, ByVal Gallary As RibbonGallery)
        Dim st As Word.Style, listTableStyle As New List(Of String)
        For Each st In doc.Styles
            If st.Type = WdStyleType.wdStyleTypeTable Then
                listTableStyle.Add(st.NameLocal)

            End If
        Next
        For Each strTableStyle As String In listTableStyle
            Dim ddi As RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem()
            With ddi
                .Label = strTableStyle
            End With
            Gallary.Items.Add(ddi)
        Next
    End Sub
    Private Sub Gallery1_Click(sender As Object, e As RibbonControlEventArgs) Handles Gallery1.Click
        F_TableStyle = Gallery1.SelectedItem.Label
        Gallery1.Label = F_TableStyle
    End Sub

    '为图片添加边框
    Private Sub Btn_AddBoarder_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_AddBoarder.Click
        Call AddBoadersForInlineshapes()
    End Sub

    '规范表格格式
    Private Sub Btn_TableFormat_Click(sender As Object, e As RibbonControlEventArgs) Handles Btn_TableFormat.Click
        Dim blnDeleteShape As Boolean = Me.CheckBox_DeleteInlineshapes.Checked
        Call TableFormat(TableStyle:=F_TableStyle, blnDeleteShapes:=blnDeleteShape)
    End Sub

    '设置超链接
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_SetHyperlinks.Click
        Call SetHyperLink()
    End Sub

    '清理文本格式
    Private Sub Button_ClearTextFormat_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_ClearTextFormat.Click
        Call ClearTextFormat()
    End Sub
#End Region

    ''' <summary>
    ''' 提取Word中的图表中的数据
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>如果将Excel中的Chart粘贴进Word，而且是以链接的形式粘贴的。在后期操作中，此Chart所链接的源Excel文件丢失，此时在Word中便不能直接提取到Excel中的数据了。</remarks>
    Private Sub ExtractDataFromWordChart(sender As Object, e As RibbonControlEventArgs) Handles btn_ExtractDataFromWordChart.Click
        Dim cht As Word.Chart = Nothing
        Dim sele As Word.Selection = App.Selection
        '先查看文档中有没有InlineShape类型的Chart
        Dim ilshps As InlineShapes, ilshp As InlineShape
        ilshps = sele.InlineShapes
        For Each ilshp In ilshps
            If ilshp.HasChart Then
                cht = ilshp.Chart
                Exit For
            End If
        Next
        '再查看文档中有没有Shape类型的Chart（即不是嵌入式图形的Chart，而是浮动式图形）
        If cht Is Nothing Then
            Dim shps As Word.ShapeRange, shp As Word.Shape
            shps = sele.ShapeRange
            For Each shp In shps
                If shp.HasChart Then
                    cht = shp.Chart
                    Exit For
                End If
            Next
        End If
        '对Chart中的数据进行提取
        If cht IsNot Nothing Then
            '
            Dim desktopPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory, Environment.SpecialFolderOption.None)
            Dim ExcelFilePath = Path.Combine({desktopPath, "Word 图表数据.xlsx"})   '用来保存数据的Excel工作簿的路径。
            Dim ExcelApp As Excel.Application, Wkbk As Excel.Workbook, sht As Excel.Worksheet
            Dim blnExcelFileExists As Boolean = False  '此Excel工作簿是否存在
            If File.Exists(ExcelFilePath) Then
                blnExcelFileExists = True
                ' 直接打开外部的文档
                Wkbk = GetObject(ExcelFilePath)  ' 打开一个Excel文档，以保存Word图表中的数据
                ExcelApp = Wkbk.Application
            Else
                ' 先创建一个Excel进程，然后再在其中添加一个工作簿。
                ExcelApp = New Excel.Application
                Wkbk = ExcelApp.Workbooks.Add()
            End If
            sht = Wkbk.Worksheets(1) ' 用工作簿中的第一个工作表来存放数据。
            sht.UsedRange.Value = Nothing
            '
            Dim seriesColl As Word.SeriesCollection = cht.SeriesCollection '这里不能定义其为Excel.SeriesCollection
            Dim Chartseries As Word.Series
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
            If blnExcelFileExists Then
                Wkbk.Save()
            Else
                Wkbk.SaveAs(Filename:=ExcelFilePath)
            End If

            sht.Activate()
            With ExcelApp
                ExcelApp.Windows(Wkbk.Name).Visible = True '取消窗口的隐藏
                ExcelApp.Windows(Wkbk.Name).Activate()
                ExcelApp.Visible = True
                If .WindowState = Excel.XlWindowState.xlMinimized Then
                    .WindowState = Excel.XlWindowState.xlNormal
                End If
            End With
        Else
            MessageBox.Show("此Word文档中没有可以进行数据提取的图表")
        End If
    End Sub

#Region "   ---   删除表格条目"

    ''' <summary>
    ''' 删除表格中的特征行
    ''' </summary>
    ''' <remarks>如果选择的区域中，某一行包含指定的标志字符，则将此行删除。
    ''' 如果选择了一个表格中的多行，则在这些行中进行检索；
    ''' 如果选择了表格中的某一个单元格，则在这一个表格的所有行中进行检索；
    ''' 这如果选择了多个表格，则在多个表格中进行检索。</remarks>
    Private Sub btnDeleteRow_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeleteRow.Click
        Dim VerifiedString As String = EditBox_standardString.Text
        Dim IdCol As UShort
        If Not UShort.TryParse(EditBox_Column.Text, IdCol) Then
            Exit Sub
        End If
        '
        App.ScreenUpdating = False
        Try
            Dim Sel As Selection = App.Selection
            Dim selectedRange As Range = Sel.Range
            Dim tables As Word.Tables = selectedRange.Tables
            Dim TableCount As UShort = tables.Count
            Dim table As Word.Table
            '
            If TableCount > 0 Then
                table = tables.Item(1)
                If TableCount = 1 Then  ' 有可能要从整个表格中去删除数据
                    If (selectedRange.Rows.Count = 1) AndAlso (selectedRange.Cells.Count < table.Columns.Count) Then ' 从这一个表格的所有行中执行删除操作
                        DeleteRow(table.Rows, VerifiedString, IdCol)
                    Else    ' 从选定的行中执行删除操作
                        DeleteRow(selectedRange.Rows, VerifiedString, IdCol)
                    End If
                Else   ' 从选择的多个表格的所有行中执行删除操作
                    For Each table In selectedRange.Tables
                        DeleteRow(table.Rows, VerifiedString, IdCol)
                    Next
                End If
            End If
        Finally
            App.ScreenUpdating = True
        End Try
    End Sub

    ''' <summary>
    ''' 从指定的集合中删除某些条目
    ''' </summary>
    ''' <param name="Rows">Rows集合</param>
    ''' <param name="VerifiedString">用来进行判断的字符串</param>
    ''' <param name="IdCol"></param>
    ''' <returns>此次一共删除了多少行</returns>
    ''' <remarks></remarks>
    Private Function DeleteRow(ByVal Rows As Word.Rows, ByVal VerifiedString As String, IdCol As UShort) As UInteger
        Dim str As String
        Dim deletedRows As UInteger
        For Each r As Row In Rows
            If IdCol <= r.Cells.Count Then
                str = r.Cells.Item(IdCol).Range.Text
                If str.IndexOf(VerifiedString, comparisonType:=StringComparison.OrdinalIgnoreCase) >= 0 Then
                    r.Delete()
                    deletedRows += 1
                End If
            End If
        Next
        Return deletedRows
    End Function
#End Region

#Region "   ---   代表的向前或者向后缩进"

    Private Sub Button_DeleteSapce_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_DeleteSapce.Click
        ' 要删除或者添加的字符数
        Dim SpaceCount As UShort
        Dim InsertSpace As String
        Try
            SpaceCount = CType(Me.EditBox_SpaceCount.Text, UShort)
            Dim sb As New System.Text.StringBuilder(4)
            For i = 1 To SpaceCount
                sb.Append(" ")
            Next
            InsertSpace = sb.ToString
        Catch ex As Exception
            MessageBox.Show("请先设置要添加或者删除的字符数")
            Exit Sub
        End Try
        '
        Try
            App.ScreenUpdating = False
            Dim StartIndex As Integer, EndIndex As Integer
            Dim rg As Range = App.Selection.Range
            StartIndex = rg.Start
            Dim str As String
            Dim rgPara As Range  '每一段的起始位置
            For Each para As Paragraph In rg.Paragraphs
                str = para.Range.Text
                If str.Length > SpaceCount AndAlso str.Substring(0, SpaceCount) = InsertSpace Then
                    rgPara = para.Range
                    rgPara.Collapse( WdCollapseDirection.wdCollapseStart)
                    rgPara.Delete(WdUnits.wdCharacter, SpaceCount)
                End If
                EndIndex = para.Range.End
            Next
            Doc.Range(StartIndex, EndIndex).Select()
        Catch ex As Exception
            '  MessageBox.Show("代码缩进出错！" & vbCrLf &
            '                   ex.Message & vbCrLf & ex.TargetSite.Name, "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            App.ScreenUpdating = True
        End Try

    End Sub

    Private Sub Button_AddSpace_Click(sender As Object, e As RibbonControlEventArgs) Handles Button_AddSpace.Click
        ' 要删除或者添加的字符数
        Dim InsertSpace As String
        Try
            Dim SpaceCount As UShort
            SpaceCount = CType(Me.EditBox_SpaceCount.Text, UShort)
            Dim sb As New System.Text.StringBuilder(4)
            For i = 1 To SpaceCount
                sb.Append(" ")
            Next
            InsertSpace = sb.ToString
        Catch ex As Exception
            MessageBox.Show("请先设置要添加或者删除的字符数")
            Exit Sub
        End Try
        '
        Try
            App.ScreenUpdating = False
            Dim StartIndex As Integer, EndIndex As Integer
            Dim rg As Range = App.Selection.Range
            StartIndex = rg.Start
            Dim rgPara As Range  '每一段的起始位置
            Dim c = rg.Paragraphs.Count
            Dim txt As String  ' 每一段的文本
            For Each para As Paragraph In rg.Paragraphs
                txt = para.Range.Text
                If txt <> Chr(13) & Chr(7) Then
                    ' 对于一个表格而言，在每一个表格的末尾，都有一个表示结尾的段落。此段落中有两个字符，所对应的ASCII码分别为13和7。
                    rgPara = para.Range
                    rgPara.Collapse(WdCollapseDirection.wdCollapseStart) ' 如果Start或End只指定一个的话，那么另一个并不会与指定了的那一个相同的。    rgPara = Doc.Range(para.Range.Start)
                    rgPara.InsertAfter(InsertSpace)
                End If
                EndIndex = para.Range.End
            Next
            Doc.Range(StartIndex, EndIndex).Select()
        Catch ex As Exception
            '  MessageBox.Show("代码缩进出错！" & vbCrLf &
            '             ex.Message & vbCrLf & ex.TargetSite.Name, "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            App.ScreenUpdating = True
        End Try
    End Sub
#End Region


#Region "  ---  子方法"

    ''' <summary>
    ''' 嵌入式图片加边框
    ''' </summary>
    ''' <param name="ParagraphStyle">此图片所在段落的段落样式</param>
    ''' <remarks></remarks>
    Sub AddBoadersForInlineshapes(Optional ByVal ParagraphStyle As String = "图片")
        Dim selection As Selection = App.Selection
        Dim picCount As Integer
        '选中区域中嵌入式图片的张数
        picCount = selection.InlineShapes.Count
        If picCount = 0 Then
            MessageBox.Show("没有发现嵌入式图片", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            App.ScreenUpdating = False
            '
            Dim Pic As InlineShape
            For Each Pic In selection.Range.InlineShapes
                '显示出图片的边框来，不然下面的设置边框线宽就会报错
                '用下面的Enable语句将图片的四个边框同时显示出来
                Pic.Borders.Enable = True
                'To remove all the borders from an object, set the Enable property to False.
                '也可以用pic.Borders(wdBorderLeft).visible = True将图片的四条边依次显示出来。
                '而对于一般的图片（即jpg等图片，而不是像AutocAD、Visio等嵌入式的对象），
                '只要设置了任意一条边的visible为true，则四条边都会同时显示出来。
                Dim rg As Range
                rg = Pic.Range
                Try
                    rg.ParagraphFormat.Style = ParagraphStyle   ' 这张图所在段落的样式为"图片"
                Catch ex As Exception
                    MessageBox.Show("请先向文档中添加样式""图片""", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try

                '对于表格中的图片，如果单元格中仅仅只有这一张图片的话，下面的添加边框的代码会失效。
                '此时要先在单元格的图片后面插入一个字符，然后添加边框，最后将字符删除。
                rg.Collapse(WdCollapseDirection.wdCollapseEnd)
                rg.InsertAfter(" ")
                '下面设置图片边框的线宽；这一定要在图片有边框时才可用，不然会报错。
                With Pic
                    .Select()
                    Dim BorderSide As Border
                    For Each BorderSide In .Borders
                        With BorderSide
                            .LineStyle = WdLineStyle.wdLineStyleSingle '边框线型wdLineStyleNone表示无边框
                            .LineWidth = WdLineWidth.wdLineWidth025pt     ' 边框线宽
                            .Color = WdColor.wdColorBlack    ' 边框颜色
                        End With
                    Next        '下一个边框

                    '设置图片的大小
                    .ScaleHeight = 100
                    .ScaleWidth = 100
                End With
                With rg
                    .Collapse(WdCollapseDirection.wdCollapseStart)
                    .Delete(Unit:=WdUnits.wdCharacter, Count:=1)
                End With
            Next    '下一张图片
            '
            selection.Collapse()
            selection.MoveRight()
            App.ScreenRefresh()
        End If
        App.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' 规范表格，而且删除表格中的嵌入式图片
    ''' </summary>
    ''' <param name="TableStyle">要应用的表格样式</param>
    ''' <param name="ParagraphFormat">表格中的段落样式</param>
    ''' <param name="blnDeleteShapes">是否要删除表格中的图片，包括嵌入式或非嵌入式图片。</param>
    ''' <remarks></remarks>
    Sub TableFormat(Optional ByVal TableStyle As String = "zengfy表格-上下总分型1", _
                    Optional ByVal ParagraphFormat As String = "表格内容置顶", _
                    Optional ByVal blnDeleteShapes As Boolean = False)
        Dim Selection = App.Selection

        If Selection.Tables.Count > 0 Then
            '定位表格
            Dim Tb As Table, rg As Range
            For Each Tb In Selection.Range.Tables
                rg = Tb.Range
                App = Tb.Application
                '
                App.ScreenUpdating = False

                '调整表格尺寸
                With Tb
                    .AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent)
                    .AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
                End With

                '清除表格中的超链接
                Dim hps As Hyperlinks
                hps = rg.Hyperlinks

                Dim hpCount As Integer
                hpCount = hps.Count
                For i = 1 To hpCount
                    hps(1).Delete()
                Next

                '将手动换行符删除
                With Tb.Range.Find
                    .ClearFormatting()
                    .Replacement.ClearFormatting()
                    .Text = "^l"
                    .Replacement.Text = ""
                    .Execute(Replace:=WdReplace.wdReplaceAll)
                End With

                '删除表格中的乱码空格
                With Tb.Range.Find
                    .ClearFormatting()
                    .Replacement.ClearFormatting()
                    .Text = " "
                    .Replacement.Text = " "
                    .Execute(Replace:=WdReplace.wdReplaceAll)
                End With

                '删除表格中的嵌入式图片
                If blnDeleteShapes Then
                    Dim inlineshps As InlineShapes, Count As Integer, inlineShp As InlineShape
                    inlineshps = Tb.Range.InlineShapes
                    Count = inlineshps.Count
                    For i = Count To 1 Step -1
                        inlineShp = inlineshps.Item(i)
                        inlineShp.Delete()
                    Next
                    '删除表格中的图片
                    Dim shps As ShapeRange, shp As Shape
                    shps = Tb.Range.ShapeRange
                    Count = shps.Count
                    For i = Count To 1 Step -1
                        shp = shps.Item(i)
                        shp.Delete()
                    Next
                End If

                '清除表格中的格式设置
                rg.Select()
                Selection.ClearFormatting()

                ' ----- 设置表格样式与表格中的段落样式
                Try        '设置表格样式
                    Tb.Style = TableStyle
                Catch ex As Exception

                End Try
                Try     '设置表格中的段落样式
                    rg.ParagraphFormat.Style = ParagraphFormat
                Catch ex As Exception

                End Try
            Next

            '取消选择并刷新界面
            Selection.Collapse()
            App.ScreenRefresh()
            App.ScreenUpdating = True
        Else
            MessageBox.Show("请至少选择一个表格。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' 设置超链接
    ''' </summary>
    ''' <remarks>此方法的要求是文本的排布格式要求：选择的段落格式必须是：
    ''' 第一段为网页标题，第二段为网址；第三段为网页标题，第四段为网址……，
    ''' 而且其中不能有空行，也不能选择空行</remarks>
    Sub SetHyperLink()
        On Error Resume Next
        Dim Selection = App.Selection
        Dim rg As Range, Prs As Paragraphs
        rg = Selection.Range
        Prs = rg.Paragraphs
        '
        Dim i As Integer
        Dim rgText As Range, rgURL As Range
        For i = Prs.Count To 1 Step -2
            '索引标题段落
            rgText = Prs.Item(i - 1).Range
            '去掉末尾的回车符
            rgText.MoveEnd(Unit:=WdUnits.wdCharacter, Count:=-1)

            '索引网址段落并得到其文本
            rgURL = Prs.Item(i).Range
            Doc.Hyperlinks.Add(Anchor:=rgText, Address:=rgURL.Text)

            '删除网址段落
            rgURL.Select()
            Selection.Delete()
        Next
    End Sub

    ''' <summary>
    ''' 清理文本的格式
    ''' </summary>
    ''' <param name="ParagraphStyle"></param>
    ''' <remarks>具体过程有：删除乱码空格、将手动换行符替换为回车、设置嵌入式图片的段落样式</remarks>
    Private Sub ClearTextFormat(Optional ByVal ParagraphStyle As String = "图片")
        On Error Resume Next
        Dim Selection = App.Selection
        Dim Sln As Selection, rg As Range
        Sln = Selection
        rg = Sln.Range

        '删除乱码空格
        With rg.Find
            .ClearFormatting()
            .Replacement.ClearFormatting()
            .Text = " "
            .Replacement.Text = " "
            .Execute(Replace:=WdReplace.wdReplaceAll)
        End With

        '将手动换行符替换为回车
        With rg.Find
            .ClearFormatting()
            .Replacement.ClearFormatting()
            .Text = "^l"
            .Replacement.Text = "^p"
            .Execute(Replace:=WdReplace.wdReplaceAll)
        End With

        Dim inlineShp As InlineShape, rgShp As Range
        For Each inlineShp In rg.InlineShapes
            inlineShp.Range.ParagraphFormat.Style = ParagraphStyle
        Next
    End Sub
#End Region


End Class
