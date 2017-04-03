using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace eZwd.eZwd_API
{
    /// <summary> 对 <see cref="Range"/> 对象进行一些通用性的操作</summary>
    public static class RangeUtils
    {

        /// <summary>
        /// 显示 Range 对象的范围
        /// </summary>
        /// <param name="rg"></param>
        public static void ShowRange(params Range[] rgs)
        {
            StringBuilder sb = new StringBuilder();
            foreach (Range rg in rgs)
            {
                if (rg != null)
                {
                    string t = rg.Text;
                    int charactorsCount = t?.Length ?? 0;
                    sb.AppendLine($"Start : {rg.Start}\t End : {rg.End}\t 字符数 :\t{charactorsCount}");
                }
            }
            MessageBox.Show(sb.ToString());
        }

        /// <summary>
        /// 将指定range范围内的特定字符串进行字符替换（不修改文字样式）
        /// </summary>
        /// <param name="rg"></param>
        /// <param name="src">要查找的字符，支持手动换行符“^l”等特殊字符</param>
        /// <param name="replace">要替换的字符，支持段落标记“^p”等特殊字符</param>
        /// <param name="matchCase"></param>
        /// <returns>returns True if the find operation is successful.</returns>
        /// <remarks>搜索或替换完成后，此Range对象的范围不会发生改变</remarks>
        public static bool ReplaceCharactors(Range rg, string src, string replace, bool matchCase = false)
        {
            Find fd = rg.Find;

            fd.ClearFormatting(); // 先清除以前的搜索选项
            fd.Text = src;

            fd.Wrap = WdFindWrap.wdFindStop;
            fd.MatchWholeWord = false;
            fd.MatchCase = matchCase;
            //
            fd.Replacement.Text = replace;
            fd.Replacement.ClearFormatting();
            //
            return fd.Execute(Replace: WdReplace.wdReplaceAll, Forward: true);
        }
    }
}
