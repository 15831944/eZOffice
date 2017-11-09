using System;
using System.Collections.Generic;
using DllActivator;
using eZstd.Enumerable;
using eZstd.Mathematics;
using eZx.AddinManager;
using eZx.RibbonHandler;
using eZx_API.Entities;
using Microsoft.Office.Interop.Excel;

namespace eZx.Debug
{
    /// <param name="excelApp"></param>
    /// <returns>如果要取消操作（即将事务 Abort 掉），则返回 false，如果要提交事务，则返回 true </returns>
    public delegate ExternalCommandResult ExternalCommand(Application excelApp);

    class AddinManagerDebuger
    {

        /// <summary>
        /// 用于 AddinManager 对代码进行调试
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="excelApp"></param>
        /// <param name="errorMessage"></param>
        /// <param name="errorRange"></param>
        /// <returns></returns>
        public static ExternalCommandResult DebugInAddinManager(ExternalCommand cmd, Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            DllActivator_eZx dat = new DllActivator_eZx();
            dat.ActivateReferences();
            try
            {
                var res = cmd(excelApp);
                switch (res)
                {

                }
                return res;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }


        /// <summary>
        /// 用于在调试完成后，最终在 Excel 的 UI 界面中，通过界面中的按钮执行相关操作
        /// </summary>
        /// <param name="cmd"></param>
        /// <param name="excelApp"></param>
        /// <param name="errorMessage"></param>
        /// <param name="errorRange"></param>
        /// <returns></returns>
        public static ExternalCommandResult ExecuteInRibbon(ExternalCommand cmd, Application excelApp, ref string errorMessage, ref Range errorRange)
        {
            //DllActivator_eZx dat = new DllActivator_eZx();
            //dat.ActivateReferences();
            try
            {
                var res = cmd(excelApp);
                switch (res)
                {

                }
                return res;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message + "\r\n\r\n" + ex.StackTrace;
                return ExternalCommandResult.Failed;
            }
            finally
            {
                excelApp.ScreenUpdating = true;
            }
        }
    }
}