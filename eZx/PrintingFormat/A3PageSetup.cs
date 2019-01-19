using eZx.AddinManager;
using eZx.Debug;
using Microsoft.Office.Interop.Excel;

namespace eZx.PrintingFormat
{
    /// <summary> A3页面打印设置 </summary>
    [EcDescription(CommandDescription)]
    class A3PageSetup : IExcelExCommand
    {
        #region --- 命令设计

        private const string CommandDescription = @"设置 A3 表格的打印格式";

        public ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            var s = new A3PageSetup();
            return AddinManagerDebuger.DebugInAddinManager(s.SetupA3Page,
                excelApp, ref errorMessage, ref errorRange);
        }

        #endregion

        private Application _excelApp;

        /// <summary> </summary>
        public ExternalCommandResult SetupA3Page(Application excelApp)
        {
            _excelApp = excelApp;
            excelApp.ScreenUpdating = false;
            var wkbk = excelApp.ActiveWorkbook;
            // 设置工作簿的“常规”样式，以确定单元格的行高
            SetupWkbkNormalStyle(wkbk);

            // 设置表格的打印样式
            var sht = excelApp.ActiveSheet as Worksheet;
            SetupSheetForPrint(sht.PageSetup, setRightHeader: true);
            return ExternalCommandResult.Succeeded;
        }

        private void SetupWkbkNormalStyle(Workbook wkbk)
        {
            var style = wkbk.Styles["Normal"];
            var font = style.Font;
            font.Name = "宋体";
            font.Size = 12;
            font.Bold = false;
            font.Italic = false;
            font.Underline = XlUnderlineStyle.xlUnderlineStyleNone;
            font.Strikethrough = false;
        }

        /// <summary> 设置 A3 表格的打印格式 </summary>
        private void SetupSheetForPrint(PageSetup ps, bool setRightHeader)
        {
            // 页边距
            ps.LeftMargin = _excelApp.CentimetersToPoints(2.5);
            ps.RightMargin = _excelApp.CentimetersToPoints(1.5);
            ps.TopMargin = _excelApp.CentimetersToPoints(1.5);
            ps.BottomMargin = _excelApp.CentimetersToPoints(1.8);

            // 页眉页脚定位
            ps.HeaderMargin = _excelApp.CentimetersToPoints(3.4);
            ps.FooterMargin = _excelApp.CentimetersToPoints(1.3);
            ps.AlignMarginsHeaderFooter = true; // 页眉页脚的内容在左右方向上与页边距对齐

            // 页眉内容
            ps.LeftHeader = "";
            ps.CenterHeader = "";
            ps.RightHeader = setRightHeader ? @"&""宋体,常规""&""Times New Roman,常规""&12第  &P  页      共  &N  页" : "";
            // 如果是 n+1 页，则为 @"&""宋体,常规""&""Times New Roman,常规""&12第  &P  页      共  &N+1  页"


            // 页脚内容
            ps.LeftFooter = new string(' ', 10) + @"&""宋体,常规""&12编制：";
            ps.CenterFooter = @"&""宋体,常规""&12校核：" + new string(' ', 20);
            ps.RightFooter = @"&""宋体,常规""&12审查：" + new string(' ', 30);

            // 正文内容在纸张中是否居中
            ps.CenterHorizontally = false;
            ps.CenterVertically = false;

            // 其他重要默认设置
            ps.BlackAndWhite = false;
            ps.PaperSize = XlPaperSize.xlPaperA3; // 纸张大小
            ps.Orientation = XlPageOrientation.xlLandscape; // 横向
            ps.Zoom = 100; // 无缩放

            //
            ps.PrintGridlines = false;
            ps.PrintComments = XlPrintLocation.xlPrintNoComments;
            ps.PrintQuality = 600;

            //
            ps.Draft = false;
            ps.FirstPageNumber = (int)Constants.xlAutomatic;
            ps.Order = XlOrder.xlDownThenOver;
            ps.PrintHeadings = false;
            ps.PrintErrors = XlPrintErrors.xlPrintErrorsDisplayed;
            ps.OddAndEvenPagesHeaderFooter = false;
            ps.DifferentFirstPageHeaderFooter = false;
            ps.ScaleWithDocHeaderFooter = true;
            ps.EvenPage.LeftHeader.Text = "";
            ps.EvenPage.CenterHeader.Text = "";
            ps.EvenPage.RightHeader.Text = "";
            ps.EvenPage.LeftFooter.Text = "";
            ps.EvenPage.CenterFooter.Text = "";
            ps.EvenPage.RightFooter.Text = "";
            ps.FirstPage.LeftHeader.Text = "";
            ps.FirstPage.CenterHeader.Text = "";
            ps.FirstPage.RightHeader.Text = "";
            ps.FirstPage.LeftFooter.Text = "";
            ps.FirstPage.CenterFooter.Text = "";
            ps.FirstPage.RightFooter.Text = "";
        }
    }
}