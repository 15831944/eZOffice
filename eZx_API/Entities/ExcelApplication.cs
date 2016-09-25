using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace eZx_API.Entities
{
    class ExcelApplication
    {
        #region ---   Properties

        public Application Application { get; }
        //
        private Workbook _activeWorkbook;
        public Workbook ActiveWorkbook { get { return _activeWorkbook; } }
        //
        private Worksheet _activeWorksheet;
        public Worksheet ActiveWorksheet { get { return _activeWorksheet; } }

        #endregion
        
        #region ---   构造函数

        /// <summary> 以隐藏的方式在后台开启一个新的Excel程序，并在其中打开指定的Excel工作簿 </summary>
        /// <param name="excelApp"> 要打开的工作簿的绝对路径 </param>
        /// <param name="visible"> 新创建的Excel程序是否要可见 </param>
        public ExcelApplication(Application excelApp, bool visible = false)
        {
            // Application
            Application = excelApp;
            Application.Visible = visible;

            // ActiveWorkbook
            var wkbks = excelApp.Workbooks;
            _activeWorkbook = wkbks.Count == 0 ? wkbks.Add() : excelApp.ActiveWorkbook;

            // ActiveWorksheet
            var shts = ActiveWorkbook.Worksheets;
            _activeWorksheet = shts.Count == 0
                ? shts.Add() as Worksheet
                : ActiveWorkbook.ActiveSheet as Worksheet;
        }

        /// <summary> 以隐藏的方式在后台开启一个新的Excel程序，并在其中打开指定的Excel工作簿 </summary>
        /// <param name="workbookPath"> 要打开的工作簿的绝对路径 </param>
        /// <param name="readOnly"> 是否要以只读的方式打开工作簿 </param>
        /// <param name="visible"> 新创建的Excel程序是否要可见 </param>
        public ExcelApplication(string workbookPath, bool readOnly = true, bool visible = false)
        {
            // Application
            var app = new Application { Visible = visible };
            Application = app;

            // ActiveWorkbook
            _activeWorkbook = app.Workbooks.Open(workbookPath, ReadOnly: readOnly);
            _activeWorkbook.Activate();

            // ActiveWorksheet
            var shts = ActiveWorkbook.Worksheets;
            _activeWorksheet = shts.Count == 0
                ? shts.Add() as Worksheet
                : ActiveWorkbook.ActiveSheet as Worksheet;
        }

        /// <summary> 以隐藏的方式在后台开启一个新的Excel程序，并在其中打开指定的Excel工作簿 </summary>
        /// <param name="workbookPath"> 要打开的工作簿的绝对路径 </param>
        /// <param name="sheetName"> 指定要打开的工作表的名称 </param>
        /// <param name="readOnly"> 是否要以只读的方式打开工作簿 </param>
        /// <param name="visible"> 新创建的Excel程序是否要可见 </param>
        public ExcelApplication(string workbookPath, string sheetName, bool readOnly = true, bool visible = false)
        {
            // Application
            var app = new Application { Visible = visible };
            Application = app;

            // ActiveWorkbook
            _activeWorkbook = app.Workbooks.Open(workbookPath, ReadOnly: readOnly);
            _activeWorkbook.Activate();

            // ActiveWorksheet
            _activeWorksheet = null;
            foreach (Worksheet sht in ActiveWorkbook.Worksheets)
            {
                if (string.Compare(sht.Name, sheetName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    _activeWorksheet = sht;
                    break;
                }
            }
            if (_activeWorksheet == null)
            {
                throw new NullReferenceException("未找到指定名称的工作表");
            }
            else
            {
                _activeWorksheet.Activate();
            }
        }

        /// <summary> 以隐藏的方式在后台开启一个新的Excel程序，并在其中打开指定的Excel工作簿 </summary>
        /// <param name="workbookPath"> 要打开的工作簿的绝对路径 </param>
        /// <param name="sheetIndex"> 指定要打开的工作表的在集合中的下标，第一个工作表的下标为1 </param>
        /// <param name="readOnly"> 是否要以只读的方式打开工作簿 </param>
        /// <param name="visible"> 新创建的Excel程序是否要可见 </param>
        public ExcelApplication(string workbookPath, int sheetIndex, bool readOnly = true, bool visible = false)
        {
            // Application
            var app = new Application { Visible = visible };
            Application = app;

            // ActiveWorkbook
            _activeWorkbook = app.Workbooks.Open(workbookPath, ReadOnly: readOnly);
            _activeWorkbook.Activate();

            // ActiveWorksheet
            _activeWorksheet = ActiveWorkbook.Worksheets.Item[sheetIndex] as Worksheet;
            _activeWorksheet.Activate();
        }

        #endregion

        /// <summary> 关闭 Excel Application 以及其中的所有工作簿，确保不留下残余进程 </summary>
        /// <param name="saveChanges"> 在关闭 Application 中的工作簿时，是否要保存对工作簿的修改 </param>
        /// <returns> 如果关闭成功，则返回true，如果关闭不完全成功，则返回false。 </returns>
        public bool SafeQuit(bool saveChanges = false)
        {

            foreach (Workbook wkbk in Application.Workbooks)
            {
                if (wkbk.ReadOnly)
                {
                    wkbk.Close(false);
                }
                else
                {
                    wkbk.Close(saveChanges);
                }
            }

            Application.Quit();

            return true;
        }

    }
}