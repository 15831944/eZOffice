using eZx.AddinManager;
using Microsoft.Office.Interop.Excel;

namespace eZx.RibbonHandler
{
    public class SheetLocker
    {
        /// <summary> 密码 </summary>
        private const string _passWord = @"JiaoTong121";
        public bool OperateOnAllSheets { get; set; }

        /// <summary> 锁定表格 </summary>
        public ExternalCommandResult LockSheet(Application app)
        {
            if (OperateOnAllSheets)
            {
                foreach (var sht in app.Worksheets)
                {
                    if (sht is Worksheet)
                    {
                        LockSheet(sht as Worksheet);
                    }
                }
            }
            else
            {
                var sht = app.ActiveSheet;
                LockSheet(sht);
            }
            return ExternalCommandResult.Succeeded;
        }

        /// <summary> 锁定表格 </summary>
        public ExternalCommandResult UnLockSheet(Application app)
        {
            if (OperateOnAllSheets)
            {
                foreach (var sht in app.Worksheets)
                {
                    if (sht is Worksheet)
                    {
                        UnLockSheet(sht as Worksheet);
                    }
                }
            }
            else
            {
                var sht = app.ActiveSheet;
                UnLockSheet(sht);
            }
            return ExternalCommandResult.Succeeded;
        }

        private void LockSheet(Worksheet sht)
        {
            sht.Protect(
                Contents: true,
                AllowFormattingCells: true,
                AllowInsertingRows: true,
                AllowInsertingHyperlinks: true,
                AllowDeletingRows: true,
                AllowSorting: true,
                Password: _passWord,
                AllowDeletingColumns: false,
                AllowInsertingColumns: false,
                AllowUsingPivotTables: false,
                DrawingObjects: false,
                Scenarios: true);
        }

        private void UnLockSheet(Worksheet sht)
        {
            sht.Unprotect(Password: _passWord);
        }
    }
}
