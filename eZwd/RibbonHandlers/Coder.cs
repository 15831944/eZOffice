using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eZwd.eZwd_API;
using Microsoft.Office.Interop.Word;

namespace eZwd.RibbonHandlers
{
    /// <summary>
    /// 与代码编辑相关的操作
    /// </summary>
    public class Coder
    {
        /// <summary> 将从Pycharm中复制到word中的代码进行格式化</summary>
        /// <param name="wdApp"></param>
        public static void FormatCodeFromIDE(Application wdApp)
        {

            Document doc = wdApp.ActiveDocument;

            Selection sel = wdApp.Selection;
            Range rg = sel.Range;
            if (rg != null)
            {
                try
                {
                    wdApp.ScreenUpdating = false;


                    // 1. change the font size
                    rg.Font.Size = 12;
                    rg.Font.Name = "Times New Roman";

                    // 2. change the shadow color to none
                    var format = rg.ParagraphFormat;
                    Shading shade = format.Shading;
                    shade.Texture = WdTextureIndex.wdTextureNone;
                    shade.ForegroundPatternColor = WdColor.wdColorAutomatic;
                    shade.BackgroundPatternColor = WdColor.wdColorAutomatic;

                    // 3. clear all tabs
                    format.TabStops.ClearAll();

                    // 4. unBold the range
                    rg.Font.Bold = 0;

                    foreach (Table tb in rg.Tables)
                    {
                        // 5. change  the talbe style 
                        tb.set_Style("zengfy表格-代码");

                        foreach (Row row in tb.Rows)
                        {
                            // 6. change the indent if the code is in a table
                            var rowFormat = row.Range.ParagraphFormat;
                            rowFormat.SpaceBeforeAuto = 0;
                            rowFormat.SpaceAfterAuto = 0;
                            rowFormat.FirstLineIndent = wdApp.CentimetersToPoints(0);
                        }
                    }

                    // 7. replace the charactors
                    RangeUtils.ReplaceCharactors(rg, "^l", "^p");

                    //
                    rg.HighlightColorIndex = WdColorIndex.wdNoHighlight;

                }
                finally
                {
                    wdApp.ScreenUpdating = true;
                }

            }
        }

    }
}
