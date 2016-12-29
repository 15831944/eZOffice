using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using eZstd.Miscellaneous;
using eZwd.ExternalCommand;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace eZwd.Debug
{
    /// <summary> 查看选择的区域的Range的范围 </summary>
    class Ec_RangeIndex : IExternalCommand
    {
        public ExternalCommandResult Execute(Microsoft.Office.Interop.Word.Application wdApp, ref string errorMessage, ref object errorObj)
        {
            try
            {
                GenerateEmbededField(wdApp);
                return ExternalCommandResult.Succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
        }

        /// <summary> 生成一个嵌套的域代码： { quote "一九一一年一月{ STYLEREF 1 \s }日" \@"D" }</summary>
        /// <param name="wdApp"></param>
        private static void GenerateEmbededField(Application wdApp)
        {
            Document doc = wdApp.ActiveDocument;
            Range rg = doc.Range();

            // ----------------------- 添加父域 -----------------------------
            // Add 方法后，Word会自动在{}域代码中的前后各补一个空格

            int fieldStart = 10;
            string quote = @"quote ""一九一一年一月日"" \@""D""";

            Field parentField = rg.Fields.Add(Range: doc.Range(fieldStart, fieldStart + 3),
                          // 如果Range参数表示一个范围，则生成的域会将此范围中的内容覆盖。而且此Range范围不受rg对象的任何影响
                          Type: WdFieldType.wdFieldEmpty,
                          Text: quote,
                          PreserveFormatting: false);
            // 如果PreserveFormatting属性值为 True，则更新时保留域所应用的格式。即会在域代码的后面附加上：\* MERGEFORMAT

            // --------------------- 在父域中添加子域 -----------------------------
            fieldStart = parentField.Code.Start + 10;
            string styleRef = @"STYLEREF 1 \s";

            rg.Fields.Add(Range: doc.Range(fieldStart, fieldStart + 5),
                //  在父域中添加子域时，Range参数只有其End属性生效，即此时定位为 parentField.Code.Start + 15，而不会覆盖任何区域
                Type: WdFieldType.wdFieldEmpty,
                Text: styleRef,
                PreserveFormatting: false);

            Type @interface = fieldStart.GetType().GetInterface("");
        }
    }
}
