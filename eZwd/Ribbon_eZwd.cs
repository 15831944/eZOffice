using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using eZwd.RibbonHandlers;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Application = Microsoft.Office.Interop.Word.Application;
using Border = Microsoft.Office.Interop.Word.Border;
using Chart = Microsoft.Office.Interop.Word.Chart;
using Hyperlinks = Microsoft.Office.Interop.Word.Hyperlinks;
using Office = Microsoft.Office.Core;
using Range = Microsoft.Office.Interop.Word.Range;
using Series = Microsoft.Office.Interop.Word.Series;
using SeriesCollection = Microsoft.Office.Interop.Word.SeriesCollection;
using Shape = Microsoft.Office.Interop.Word.Shape;
using ShapeRange = Microsoft.Office.Interop.Word.ShapeRange;
using Style = Microsoft.Office.Interop.Word.Style;

namespace eZwd
{
    internal partial class Ribbon_eZwd
    {
        #region   ---  Properties

        #endregion

        #region   ---  Fields

        private Application _app;

        /// <summary> 当前正在运行的Word程序中的活动Word文档
        /// </summary> <remarks></remarks>
        private Document _activeDoc;

        /// <summary>
        /// 进行表格规范化时所使用的表格样式
        /// </summary>
        /// <remarks>
        /// 注意：在为内容指定样式（比如为段落指定段落样式或者为表格指定表格样式）时，
        /// 如果指定的样式不存在或者为段落指定了表格样式等时，程序会继续正常执行，也不会跳过后面的语句，
        /// 只是就相当于没有执行这一行。</remarks>
        private string F_TableStyle = "zengfy表格-上下总分型1";

        #endregion

        #region   ---  构造函数与窗体的加载、打开与关闭

        private void Ribbon_eZwd_Load(Object sender, RibbonUIEventArgs e)
        {
            _app = Globals.ThisAddIn.Application;
            _app.DocumentChange += AppOnDocumentChange;
        }

        private void AppOnDocumentChange()
        {
            if (_app.Documents.Count > 0)
            {
                _activeDoc = _app.ActiveDocument;
                ListStyles(_activeDoc, this.Gallery1);
            }
        }

        #endregion

        #region   ---  界面操作

        //列出与选择表格样式
        /// <summary>
        /// 列出文档中所有的表格样式
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="gallary"></param>
        /// <remarks></remarks>
        private void ListStyles(Document doc, RibbonGallery gallary)
        {
            gallary.Items.Clear();
            Style st;
            List<string> listTableStyle = new List<string>();
            foreach (Style tempLoopVar_st in doc.Styles)
            {
                st = tempLoopVar_st;
                if (st.Type == WdStyleType.wdStyleTypeTable)
                {
                    listTableStyle.Add(st.NameLocal);
                }
            }
            foreach (string strTableStyle in listTableStyle)
            {
                RibbonDropDownItem ddi = this.Factory.CreateRibbonDropDownItem();
                ddi.Label = strTableStyle;
                gallary.Items.Add(ddi);
            }
        }

        public void Gallery1_Click(object sender, RibbonControlEventArgs e)
        {
            F_TableStyle = Gallery1.SelectedItem.Label;
            Gallery1.Label = F_TableStyle;
        }

        //为图片添加边框
        public void Btn_AddBoarder_Click(object sender, RibbonControlEventArgs e)
        {
            Application wdApp = Globals.ThisAddIn.Application;
            StaticFunction.AddBoadersForInlineshapes(wdApp, pictureParagraphStyle: "图片");
        }

        //规范表格格式
        public void Btn_TableFormat_Click(object sender, RibbonControlEventArgs e)
        {
            bool blnDeleteShape = this.CheckBox_DeleteInlineshapes.Checked;
            TableFormat(TableStyle: F_TableStyle, blnDeleteShapes: blnDeleteShape);
        }

        //设置超链接
        public void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Application wdApp = Globals.ThisAddIn.Application;
            StaticFunction.SetHyperLink(wdApp);
        }

        //清理文本格式
        public void Button_ClearTextFormat_Click(object sender, RibbonControlEventArgs e)
        {
            Application wdApp = Globals.ThisAddIn.Application;
            StaticFunction.ClearTextFormat(wdApp, pictureParagraphStyle: "图片");
        }

        /// <summary>
        /// 将多个段落转换为一个段落
        /// </summary>
        /// <remarks>比如将从PDF中粘贴过来的多段文字转换为一个段落。具体操作为：将选择区域的文字中的换行符转换为空格</remarks>
        public void buttonPdfReformat_Click(object sender, RibbonControlEventArgs e)
        {
            Application wdApp = Globals.ThisAddIn.Application;
            StaticFunction.PdfReformat(wdApp);
        }

        #endregion

        /// <summary>
        /// 提取Word中的图表中的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks>如果将Excel中的Chart粘贴进Word，而且是以链接的形式粘贴的。在后期操作中，此Chart所链接的源Excel文件丢失，此时在Word中便不能直接提取到Excel中的数据了。</remarks>
        public void ExtractDataFromWordChart(object sender, RibbonControlEventArgs e)
        {
            Chart cht = null;
            Selection sele = _app.Selection;
            //先查看文档中有没有InlineShape类型的Chart
            InlineShapes ilshps = default(InlineShapes);
            InlineShape ilshp = default(InlineShape);
            ilshps = sele.InlineShapes;
            foreach (InlineShape tempLoopVar_ilshp in ilshps)
            {
                ilshp = tempLoopVar_ilshp;
                if (ilshp.HasChart == Office.MsoTriState.msoTrue)
                {
                    cht = ilshp.Chart;
                    break;
                }
            }
            //再查看文档中有没有Shape类型的Chart（即不是嵌入式图形的Chart，而是浮动式图形）
            if (cht == null)
            {
                ShapeRange shps = default(ShapeRange);
                Shape shp = default(Shape);
                shps = sele.ShapeRange;
                foreach (Shape tempLoopVar_shp in shps)
                {
                    shp = tempLoopVar_shp;
                    if (shp.HasChart == Office.MsoTriState.msoTrue)
                    {
                        cht = shp.Chart;
                        break;
                    }
                }
            }
            //对Chart中的数据进行提取
            if (cht != null)
            {
                //
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory,
                    Environment.SpecialFolderOption.None);
                var ExcelFilePath = Path.Combine(new[] { desktopPath, "Word 图表数据.xlsx" }); //用来保存数据的Excel工作簿的路径。
                Microsoft.Office.Interop.Excel.Application ExcelApp =
                    default(Microsoft.Office.Interop.Excel.Application);
                Workbook Wkbk = default(Workbook);
                Worksheet sht = default(Worksheet);
                bool blnExcelFileExists = false; //此Excel工作簿是否存在
                if (File.Exists(ExcelFilePath.ToString()))
                {
                    blnExcelFileExists = true;
                    // 直接打开外部的文档
                    Wkbk = (Workbook)Interaction.GetObject(ExcelFilePath.ToString(), null);
                    // 打开一个Excel文档，以保存Word图表中的数据
                    ExcelApp = Wkbk.Application;
                }
                else
                {
                    // 先创建一个Excel进程，然后再在其中添加一个工作簿。
                    ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Wkbk = ExcelApp.Workbooks.Add();
                }
                sht = Wkbk.Worksheets[1]; // 用工作簿中的第一个工作表来存放数据。
                sht.UsedRange.Value = null;
                //
                SeriesCollection seriesColl = (SeriesCollection)cht.SeriesCollection();
                //这里不能定义其为Excel.SeriesCollection
                Series Chartseries;
                //开始提取数据
                short col = (short)1;
                dynamic X = default(dynamic); // 这里只能将X与Y的数据类型定义为Object，不能是Object()或者Object(,)
                object Y = null;
                string Title = "";
                // 这里不能用For Each Chartseries in SeriesCollection来引用seriesCollection集合中的元素。
                for (var i = 1; i <= seriesColl.Count; i++)
                {
                    // 在VB.NET中，seriesCollection集合中的第一个元素的下标值为1。
                    Chartseries = seriesColl[i];
                    X = Chartseries.XValues;
                    Y = Chartseries.Values;
                    Title = Chartseries.Name;
                    // 将数据存入Excel表中
                    int PointsCount = Convert.ToInt32(X.Length);
                    if (PointsCount > 0)
                    {
                        sht.Cells[1, col].Value = Title;
                        sht.Range[sht.Cells[2, col], sht.Cells[PointsCount + 1, col]].Value =
                            ExcelApp.WorksheetFunction.Transpose(X);
                        sht.Range[sht.Cells[2, col + 1], sht.Cells[PointsCount + 1, col + 1]].Value =
                            ExcelApp.WorksheetFunction.Transpose(Y);
                        col = (short)(col + 3);
                    }
                }
                if (blnExcelFileExists)
                {
                    Wkbk.Save();
                }
                else
                {
                    Wkbk.SaveAs(Filename: ExcelFilePath);
                }

                sht.Activate();
                ExcelApp.Windows[Wkbk.Name].Visible = true; //取消窗口的隐藏
                ExcelApp.Windows[Wkbk.Name].Activate();
                ExcelApp.Visible = true;
                if (ExcelApp.WindowState == XlWindowState.xlMinimized)
                {
                    ExcelApp.WindowState = XlWindowState.xlNormal;
                }
            }
            else
            {
                MessageBox.Show("此Word文档中没有可以进行数据提取的图表");
            }
        }

        #region    ---   删除表格条目

        /// <summary>
        /// 删除表格中的特征行
        /// </summary>
        /// <remarks>如果选择的区域中，某一行包含指定的标志字符，则将此行删除。
        /// 如果选择了一个表格中的多行，则在这些行中进行检索；
        /// 如果选择了表格中的某一个单元格，则在这一个表格的所有行中进行检索；
        /// 这如果选择了多个表格，则在多个表格中进行检索。</remarks>
        public void btnDeleteRow_Click(object sender, RibbonControlEventArgs e)
        {
            string VerifiedString = EditBox_standardString.Text;
            UInt16 IdCol = 0;
            if (!UInt16.TryParse(EditBox_Column.Text, out IdCol))
            {
                return;
            }
            //
            _app.ScreenUpdating = false;
            try
            {
                Selection Sel = _app.Selection;
                Range selectedRange = Sel.Range;
                Tables tables = selectedRange.Tables;
                UInt16 TableCount = (UInt16)tables.Count;
                Table table = default(Table);
                //
                if (TableCount > 0)
                {
                    table = tables[1];
                    if (TableCount == 1) // 有可能要从整个表格中去删除数据
                    {
                        if ((selectedRange.Rows.Count == 1) && (selectedRange.Cells.Count < table.Columns.Count))
                        // 从这一个表格的所有行中执行删除操作
                        {
                            DeleteRow(table.Rows, VerifiedString, IdCol);
                        }
                        else // 从选定的行中执行删除操作
                        {
                            DeleteRow(selectedRange.Rows, VerifiedString, IdCol);
                        }
                    }
                    else // 从选择的多个表格的所有行中执行删除操作
                    {
                        foreach (Table tempLoopVar_table in selectedRange.Tables)
                        {
                            table = tempLoopVar_table;
                            DeleteRow(table.Rows, VerifiedString, IdCol);
                        }
                    }
                }
            }
            finally
            {
                _app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// 从指定的集合中删除某些条目
        /// </summary>
        /// <param name="Rows">Rows集合</param>
        /// <param name="VerifiedString">用来进行判断的字符串</param>
        /// <param name="IdCol"></param>
        /// <returns>此次一共删除了多少行</returns>
        /// <remarks></remarks>
        private UInt32 DeleteRow(Rows Rows, string VerifiedString, UInt16 IdCol)
        {
            string str = "";
            UInt32 deletedRows = 0;
            foreach (Row r in Rows)
            {
                if (IdCol <= r.Cells.Count)
                {
                    str = Convert.ToString(r.Cells[IdCol].Range.Text);
                    if (str.IndexOf(value: VerifiedString, comparisonType: StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        r.Delete();
                        deletedRows++;
                    }
                }
            }
            return deletedRows;
        }

        #endregion

        #region    ---   代码编辑相关操作

        /// <summary> 代码向前缩进 </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Button_DeleteSapce_Click(object sender, RibbonControlEventArgs e)
        {
            // 要删除或者添加的字符数
            int SpaceCount = 0;
            string InsertSpace = "";
            try
            {
                SpaceCount = Convert.ToUInt16(this.EditBox_SpaceCount.Text);
                StringBuilder sb = new StringBuilder(4);
                for (var i = 1; i <= SpaceCount; i++)
                {
                    sb.Append(" ");
                }
                InsertSpace = sb.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("请先设置要添加或者删除的字符数");
                return;
            }
            //
            try
            {
                _app.ScreenUpdating = false;
                int StartIndex = 0;
                int EndIndex = 0;
                Range rg = _app.Selection.Range;
                StartIndex = rg.Start;
                string str = "";
                Range rgPara = default(Range); //每一段的起始位置
                foreach (Paragraph para in rg.Paragraphs)
                {
                    str = para.Range.Text;
                    if (str.Length > SpaceCount && str.Substring(0, SpaceCount) == InsertSpace)
                    {
                        rgPara = para.Range;
                        rgPara.Collapse(Direction: WdCollapseDirection.wdCollapseStart);
                        rgPara.Delete(Unit: WdUnits.wdCharacter, Count: SpaceCount);
                    }
                    EndIndex = para.Range.End;
                }
                _activeDoc.Range(StartIndex, EndIndex).Select();
            }
            catch (Exception)
            {
                //  MessageBox.Show("代码缩进出错！" & vbCrLf &
                //                   ex.Message & vbCrLf & ex.TargetSite.Name, "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
            }
            finally
            {
                _app.ScreenUpdating = true;
            }
        }

        /// <summary> 代码向前缩进 </summary>
        public void Button_AddSpace_Click(object sender, RibbonControlEventArgs e)
        {
            // 要删除或者添加的字符数
            string InsertSpace = "";
            try
            {
                UInt16 SpaceCount = 0;
                SpaceCount = Convert.ToUInt16(this.EditBox_SpaceCount.Text);
                StringBuilder sb = new StringBuilder(4);
                for (var i = 1; i <= SpaceCount; i++)
                {
                    sb.Append(" ");
                }
                InsertSpace = sb.ToString();
            }
            catch (Exception)
            {
                MessageBox.Show("请先设置要添加或者删除的字符数");
                return;
            }
            //
            try
            {
                _app.ScreenUpdating = false;
                int StartIndex = 0;
                int EndIndex = 0;
                Range rg = _app.Selection.Range;
                StartIndex = rg.Start;
                Range rgPara = default(Range); //每一段的起始位置
                var c = rg.Paragraphs.Count;
                string txt = ""; // 每一段的文本
                foreach (Paragraph para in rg.Paragraphs)
                {
                    txt = para.Range.Text;
                    if (txt != '\r' + "\a")
                    {
                        // 对于一个表格而言，在每一个表格的末尾，都有一个表示结尾的段落。此段落中有两个字符，所对应的ASCII码分别为13和7。
                        rgPara = para.Range;
                        rgPara.Collapse(Direction: WdCollapseDirection.wdCollapseStart);
                        // 如果Start或End只指定一个的话，那么另一个并不会与指定了的那一个相同的。    rgPara = Doc.Range(para.Range.Start)
                        rgPara.InsertAfter(InsertSpace);
                    }
                    EndIndex = para.Range.End;
                }

                _activeDoc.Range(StartIndex, EndIndex).Select();
            }
            catch (Exception)
            {
                //  MessageBox.Show("代码缩进出错！" & vbCrLf &
                //             ex.Message & vbCrLf & ex.TargetSite.Name, "出错", MessageBoxButtons.OK, MessageBoxIcon.Error)
            }
            finally
            {
                _app.ScreenUpdating = true;
            }
        }


        private void button_CodeFormater_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            RibbonHandlers.Coder.FormatCodeFromIDE(app);
        }
        #endregion

        #region    ---   快速引用

        private void btn_CrossRef_Click(object sender, RibbonControlEventArgs e)
        {
            if (FastCrossReference.CitedRg == null)
            {
                var app = Globals.ThisAddIn.Application;
                FastCrossReference.SetAnchor(app);
                UI_CrossRefStarted();
            }
            else
            {
                FastCrossReference.Exit();
                UI_CrossRefEnded();
            }
        }

        private void btn_CrossRefExecute_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            bool ret = chk_CrossRefReturn.Checked;
            var succ = FastCrossReference.CrossRef(app, ret);
            if (succ)
            {
                FastCrossReference.Exit();
                UI_CrossRefEnded();
            }
        }

        private void UI_CrossRefStarted()
        {
            btn_CrossRef.Label = "取消引用";
            btn_CrossRefExecute.Enabled = true;
            chk_CrossRefReturn.Enabled = true;
        }
        private void UI_CrossRefEnded()
        {
            btn_CrossRef.Label = "锚定引用";
            btn_CrossRefExecute.Enabled = false;
            chk_CrossRefReturn.Enabled = false;
        }
        #endregion

        #region   ---  子方法

        /// <summary>
        /// 规范表格，而且删除表格中的嵌入式图片
        /// </summary>
        /// <param name="TableStyle">要应用的表格样式</param>
        /// <param name="ParagraphFormat">表格中的段落样式</param>
        /// <param name="blnDeleteShapes">是否要删除表格中的图片，包括嵌入式或非嵌入式图片。</param>
        /// <remarks></remarks>
        public void TableFormat(string TableStyle = "zengfy表格-上下总分型1", string ParagraphFormat = "表格内容置顶",
            bool blnDeleteShapes = false)
        {
            var Selection = _app.Selection;

            if (Selection.Tables.Count > 0)
            {
                //定位表格
                Table Tb = default(Table);
                Range rg = default(Range);
                foreach (Table tempLoopVar_Tb in Selection.Range.Tables)
                {
                    Tb = tempLoopVar_Tb;
                    rg = Tb.Range;
                    _app = Tb.Application;
                    //
                    _app.ScreenUpdating = false;

                    //调整表格尺寸
                    dynamic with_1 = Tb;
                    with_1.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                    with_1.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

                    //清除表格中的超链接
                    Hyperlinks hps = default(Hyperlinks);
                    hps = rg.Hyperlinks;

                    int hpCount = 0;
                    hpCount = hps.Count;
                    for (var i = 1; i <= hpCount; i++)
                    {
                        hps[1].Delete();
                    }

                    //将手动换行符删除
                    {
                        Tb.Range.Find.ClearFormatting();
                        Tb.Range.Find.Replacement.ClearFormatting();
                        Tb.Range.Find.Text = "^l";
                        Tb.Range.Find.Replacement.Text = "";
                        Tb.Range.Find.Execute(WdReplace.wdReplaceAll);
                    }
                    //删除表格中的乱码空格
                    {
                        Tb.Range.Find.ClearFormatting();
                        Tb.Range.Find.Replacement.ClearFormatting();
                        Tb.Range.Find.Text = " ";
                        Tb.Range.Find.Replacement.Text = " ";
                        Tb.Range.Find.Execute(WdReplace.wdReplaceAll);
                    }

                    //删除表格中的嵌入式图片
                    if (blnDeleteShapes)
                    {
                        InlineShapes inlineshps = default(InlineShapes);
                        int Count = 0;
                        InlineShape inlineShp = default(InlineShape);
                        inlineshps = Tb.Range.InlineShapes;
                        Count = inlineshps.Count;
                        for (var i = Count; i >= 1; i--)
                        {
                            inlineShp = inlineshps[Convert.ToInt32(i)];
                            inlineShp.Delete();
                        }
                        //删除表格中的图片
                        ShapeRange shps = default(ShapeRange);
                        Shape shp = default(Shape);
                        shps = Tb.Range.ShapeRange;
                        Count = shps.Count;
                        for (var i = Count; i >= 1; i--)
                        {
                            shp = shps[i];
                            shp.Delete();
                        }
                    }

                    //清除表格中的格式设置
                    rg.Select();
                    Selection.ClearFormatting();

                    // ----- 设置表格样式与表格中的段落样式
                    try //设置表格样式
                    {
                        Tb.set_Style(TableStyle);
                    }
                    catch (Exception)
                    {
                    }
                    try //设置表格中的段落样式
                    {
                        rg.ParagraphFormat.set_Style(ParagraphFormat);
                    }
                    catch (Exception)
                    {
                    }
                }

                //取消选择并刷新界面
                Selection.Collapse();
                _app.ScreenRefresh();
                _app.ScreenUpdating = true;
            }
            else
            {
                MessageBox.Show("请至少选择一个表格。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        #endregion

    }
}