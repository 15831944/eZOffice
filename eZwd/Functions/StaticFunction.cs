using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace eZwd.Functions
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



        /// <summary>
        /// 清理文本的格式
        /// </summary>
        /// <param name="wdApp"></param>
        /// <param name="pictureParagraphStyle">图片所在段落的段落样式</param>
        /// <remarks>具体过程有：删除乱码空格、将手动换行符替换为回车、设置嵌入式图片的段落样式</remarks>
        public static void ClearTextFormat(Application wdApp, string pictureParagraphStyle = "图片")
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
    }
}
