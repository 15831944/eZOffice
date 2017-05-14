using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.RibbonHandlers
{
    /// <summary> 快速添加交叉引用 </summary>
    public class FastCrossReference
    {
        public static Range CitedRg;

        public static void SetAnchor(Application wdApp)
        {
            var rg = wdApp.Selection.Range;
            if (rg == null) return;
            CitedRg = rg;
        }

        /// <summary>
        /// 插入交叉引用
        /// </summary>
        /// <param name="wdApp"></param>
        /// <param name="returnToCited">是否要返回到发起引用的位置</param>
        /// <returns></returns>
        public static bool CrossRef(Application wdApp, bool returnToCited)
        {
            wdApp.ScreenUpdating = false;
            bool succ = false;
            if (CitedRg == null) return true;
            var desti = wdApp.Selection.Range;
            if (desti == null) return false;
            //
            var doc = desti.Document;
            int paraIndex = FindCrossRefIndex(doc, desti);

            if (paraIndex != 0)
            {
                try
                {
                    // 插入交叉引用，插入后会将range中的所有内容替换为插入的引用项，并将其重新Collapse到range的开头。
                    // 插入标题内容
                    CitedRg.InsertCrossReference(
                        ReferenceType: WdReferenceType.wdRefTypeHeading,
                        ReferenceKind: WdReferenceKind.wdContentText,
                        ReferenceItem: paraIndex,
                        InsertAsHyperlink: true,
                        IncludePosition: false,
                        SeparateNumbers: false,
                        SeparatorString: " ");

                    // 在标题编号与标题内容之间插入一个空格。
                    CitedRg.InsertAfter(" ");
                    CitedRg.Collapse(WdCollapseDirection.wdCollapseStart);

                    // 插入标题编号
                    CitedRg.InsertCrossReference(
                        ReferenceType: WdReferenceType.wdRefTypeHeading,
                        ReferenceKind: WdReferenceKind.wdNumberFullContext,
                        ReferenceItem: paraIndex,
                        InsertAsHyperlink: true,
                        IncludePosition: false,
                        SeparateNumbers: false,
                        SeparatorString: " ");
                    succ = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + "\r\n" + @"检查被引用的大纲标题是否没有标题内容。",
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    // 比如大纲段落中只有编号，没有内容，则会出错“需要引用的内容为空”
                    // 比如大纲段落中没有编号
                    succ = false;
                }
            }
            else
            {
                MessageBox.Show(@"请选择具有编号的标题段落", @"提示", MessageBoxButtons.OK);
                succ = false;
            }

            if (returnToCited)
            {
                CitedRg.Select();
                // 将界面滚动到指定位置
                wdApp.ActiveWindow.ScrollIntoView(CitedRg);
            }
            wdApp.ScreenUpdating = true;
            return succ;
        }

        /// <summary>
        /// 被引用的目标段落在整个交叉引用的“标题”集合中的下标
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="destPara"></param>
        /// <returns></returns>
        private static int FindCrossRefIndex(Document doc, Range destPara)
        {
            var paraRg = destPara.Paragraphs[1].Range;
            // 引用段落的标志性内容，比如“1.1 测试小节”
            var paraText = $"{paraRg.ListFormat.ListString} {paraRg.Text.TrimEnd('\r')}";
            
            int paraIndex = 0;

            // 返回一个 string[] 集合，而且其第一个元素的下标为1
            var refItems = doc.Application.ActiveDocument.GetCrossReferenceItems(WdReferenceType.wdRefTypeHeading);
            var itemsArr = refItems as Array; // 必须将其转换为 Array，而不能是 string[]，否则会报错。

            //
            int ind = 0;
            foreach (string s in itemsArr)
            {
                ind += 1;
                var st = s.TrimStart(' '); // 删除编号前面的定位空格
                if (string.Equals(st, paraText))
                {
                    return ind;
                }
            }
            return paraIndex;
        }

        public static void Exit()
        {
            CitedRg = null;
        }
    }
}