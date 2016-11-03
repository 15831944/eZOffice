using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.RibbonHandlers
{
    /// <summary> 一些结构性的静态方法 </summary>
    internal static class StaticFunction
    {
        /// <summary>
        /// 将多个段落转换为一个段落
        /// </summary>
        /// <param name="wdApp"></param>
        /// <remarks>比如将从PDF中粘贴过来的多段文字转换为一个段落。具体操作为：将选择区域的文字中的换行符转换为空格</remarks>
        public static void PdfReformat(Application wdApp)
        {
            Selection sele = wdApp.Selection;
            if (sele != null && sele.Start != sele.End)
            {
                int startIndex = sele.Start;
                int endIndex = sele.End;
                int paraEndIndex = sele.Paragraphs[1].Range.End;
                // 必须包含至少一个段落
                if (paraEndIndex > endIndex)
                {
                    return;
                }

                // 进行替换
                Find fd = sele.Range.Find;
                fd.ClearFormatting();
                fd.Replacement.ClearFormatting();
                // 设置替换格式

                fd.Text = "^p";
                fd.Replacement.Text = " "; // 将换行符替换为空格
                fd.Forward = true;
                fd.Wrap = WdFindWrap.wdFindStop;
                fd.Format = false;
                fd.MatchCase = false;
                fd.MatchWholeWord = false;
                fd.MatchByte = true;
                fd.MatchWildcards = false;
                fd.MatchSoundsLike = false;
                fd.MatchAllWordForms = false;
                //
                fd.Execute(Replace: WdReplace.wdReplaceAll);

                // 删除最后面的那一个空格
                sele.MoveRight(Unit: WdUnits.wdCharacter, Count: 1);
                sele.TypeBackspace();
                sele.TypeParagraph();

                // 选择重新格式化后的这一段落
                endIndex = sele.End - 1;
                Range rg = wdApp.ActiveDocument.Range(Start: startIndex, End: endIndex);
                rg.Select();
            }
        }


        /// <summary> 设置超链接 </summary>
        /// <param name="wdApp"></param>
        /// <remarks>此方法的要求是文本的排布格式要求：选择的段落格式必须是：
        /// 第一段为网页标题，第二段为网址；第三段为网页标题，第四段为网址……，
        /// 而且其中不能有空行，也不能选择空行</remarks>
        public static void SetHyperLink(Application wdApp)
        {
            var sele = wdApp.Selection;
            Range rg = default(Range);
            Paragraphs Prs = default(Paragraphs);
            rg = sele.Range;
            Prs = rg.Paragraphs;
            //
            int i = 0;
            Range rgText = default(Range);
            Range rgURL = default(Range);
            for (i = Prs.Count; i >= 1; i -= 2)
            {
                //索引标题段落
                rgText = Prs[i - 1].Range;
                //去掉末尾的回车符
                rgText.MoveEnd(Unit: WdUnits.wdCharacter, Count: -1);

                //索引网址段落并得到其文本
                rgURL = Prs[i].Range;
                wdApp.ActiveDocument.Hyperlinks.Add(Anchor: rgText, Address: rgURL.Text);

                //删除网址段落
                rgURL.Select();
                sele.Delete();
            }
        }


        /// <summary>
        /// 清理文本的格式
        /// </summary>
        /// <param name="wdApp"></param>
        /// <param name="pictureParagraphStyle">图片所在段落的段落样式</param>
        /// <remarks>具体过程有：删除乱码空格、将手动换行符替换为回车、设置嵌入式图片的段落样式</remarks>
        public static void ClearTextFormat(Application wdApp, string pictureParagraphStyle)
        {
            //On Error Resume Next VBConversions Warning: On Error Resume Next not supported in C#
            var sele = wdApp.Selection;
            Range rg = sele.Range;

            //删除乱码空格
            rg.Find.ClearFormatting();
            rg.Find.Replacement.ClearFormatting();
            rg.Find.Text = " ";
            rg.Find.Replacement.Text = " ";
            rg.Find.Execute(Replace: WdReplace.wdReplaceAll);

            //将手动换行符替换为回车
            rg.Find.ClearFormatting();
            rg.Find.Replacement.ClearFormatting();
            rg.Find.Text = "^l";
            rg.Find.Replacement.Text = "^p";
            rg.Find.Execute(Replace: WdReplace.wdReplaceAll);

            //
            foreach (InlineShape inlineShp in rg.InlineShapes)
            {
                inlineShp.Range.ParagraphFormat.set_Style(pictureParagraphStyle);
            }
        }


        /// <summary>
        /// 嵌入式图片加边框
        /// </summary>
        /// <param name="wdApp"></param>
        /// <param name="pictureParagraphStyle">此图片所在段落的段落样式</param>
        /// <remarks></remarks>
        public static void AddBoadersForInlineshapes(Application wdApp, string pictureParagraphStyle)
        {
            Selection selection = wdApp.Selection;
            int picCount;
            //选中区域中嵌入式图片的张数
            picCount = selection.InlineShapes.Count;
            if (picCount == 0)
            {
                MessageBox.Show("没有发现嵌入式图片", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                wdApp.ScreenUpdating = false;
                //
                InlineShape Pic = default(InlineShape);
                foreach (InlineShape tempLoopVar_Pic in selection.Range.InlineShapes)
                {
                    Pic = tempLoopVar_Pic;
                    //显示出图片的边框来，不然下面的设置边框线宽就会报错
                    //用下面的Enable语句将图片的四个边框同时显示出来
                    Pic.Borders.Enable = 1; // true;
                    //To remove all the borders from an object, set the Enable property to False.
                    //也可以用pic.Borders(wdBorderLeft).visible = True将图片的四条边依次显示出来。
                    //而对于一般的图片（即jpg等图片，而不是像AutocAD、Visio等嵌入式的对象），
                    //只要设置了任意一条边的visible为true，则四条边都会同时显示出来。
                    Range rg = default(Range);
                    rg = Pic.Range;
                    try
                    {
                        // rg.ParagraphFormat.Style = ParagraphStyle; 
                        rg.ParagraphFormat.set_Style(pictureParagraphStyle); // 这张图所在段落的样式为"图片"
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("请先向文档中添加样式\"图片\"", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    //对于表格中的图片，如果单元格中仅仅只有这一张图片的话，下面的添加边框的代码会失效。
                    //此时要先在单元格的图片后面插入一个字符，然后添加边框，最后将字符删除。
                    rg.Collapse(Direction: WdCollapseDirection.wdCollapseEnd);
                    rg.InsertAfter(" ");
                    //下面设置图片边框的线宽；这一定要在图片有边框时才可用，不然会报错。
                    dynamic with_1 = Pic;
                    with_1.Select();
                    Border BorderSide = default(Border);
                    foreach (Border tempLoopVar_BorderSide in with_1.Borders)
                    {
                        BorderSide = tempLoopVar_BorderSide;
                        dynamic with_2 = BorderSide;
                        with_2.LineStyle = WdLineStyle.wdLineStyleSingle; //边框线型wdLineStyleNone表示无边框
                        with_2.LineWidth = WdLineWidth.wdLineWidth025pt; // 边框线宽
                        with_2.Color = WdColor.wdColorBlack; // 边框颜色
                    } //下一个边框

                    //设置图片的大小
                    with_1.ScaleHeight = 100;
                    with_1.ScaleWidth = 100;
                    rg.Collapse(WdCollapseDirection.wdCollapseStart);
                    rg.Delete(Unit: WdUnits.wdCharacter, Count: 1);
                } //下一张图片

                //
                selection.Collapse();
                selection.MoveRight();
                wdApp.ScreenRefresh();
            }
            wdApp.ScreenUpdating = true;
        }

    }
}
