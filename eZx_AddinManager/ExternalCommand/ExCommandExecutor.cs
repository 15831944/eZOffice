﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.AddinManager
{
    internal static class ExCommandExecutor
    {
        private static IExcelExCommand _currentExternalCommand;
        private static string _currentExternalCommandAssemblyPath;

        /// <summary> 执行当前（即上次执行过的那个）外部命令 </summary>
        /// <param name="excelApp">作为Execute()方法的输入参数的对象，表示当前的Excel Application </param>
        public static void InvokeCurrentExternalCommand(Application excelApp)
        {
            if (_currentExternalCommand != null && _currentExternalCommandAssemblyPath != null)
            {
                InvokeExternalCommand(_currentExternalCommandAssemblyPath, _currentExternalCommand, excelApp);
            }
        }

        /// <summary> 执行 CAD 的外部命令 </summary>
        /// <param name="assemblyPath">外部命令所对应的dll的绝对路径</param>
        /// <param name="externalCommand">此命令必须是实现了 IExternalCommand.Execute </param>
        /// <param name="excelApp">作为Execute()方法的输入参数的对象，表示当前的Excel Application </param>
        /// <remarks>出于调试的即时更新的考虑，这里在每一次调试外部命令时，都对最新的dll进行重新加载。</remarks>
        public static void InvokeExternalCommand(string assemblyPath, IExcelExCommand externalCommand, Application excelApp)
        {

            ExCommandExecutor.RunActiveCommand(externalCommand, assemblyPath, excelApp);
            //
            _currentExternalCommandAssemblyPath = assemblyPath;
            _currentExternalCommand = externalCommand;

        }


        /// <summary> 执行 Excel 的外部命令 </summary>
        /// <param name="externalCommandMethod">此命令必须是实现了 IExternalCommand.Execute </param>
        private static void InvokeCommand(MethodInfo externalCommandMethod, Application excelApp)
        {
            // 创建一个实例对象
            Type tp = externalCommandMethod.DeclaringType;
            object ins = tp.Assembly.CreateInstance(tp.FullName);

            //// 构造一个实例，在此构造函数中即可以设计对应的插件功能了
            //object instance = asm.CreateInstance(
            //    typeName: classFullName,
            //    ignoreCase: false,
            //    bindingAttr: BindingFlags.CreateInstance,
            //    binder: null,
            //    args: constructorArgs,
            //    culture: CultureInfo.CurrentCulture,
            //    activationAttributes: null);

            // 执行命令
            ExternalCommandResult res;

            string errorMessage = "";
            Range errorRange = null;

            //
            // 注意如果要提取 ref 或 out 类型的参数的结果，则必须将对应的参数全部放置在一个 parameters 数组中
            object[] parameters = new object[] { excelApp, "", errorRange };
            try
            {
                // 执行操作
                res = (ExternalCommandResult)externalCommandMethod.Invoke(obj: ins, parameters: parameters);

                // 提取 ref 类型 或者 out 类型的 参数
                errorMessage = parameters[1] as string;
                errorRange = parameters[2] as Range;
            }
            catch (Exception ex)
            {
                // 提取 ref 类型 或者 out 类型的 参数
                errorMessage = parameters[1] as string;
                //
                errorRange = parameters[2] as Range;
                if (string.IsNullOrEmpty(errorMessage))
                {
                    errorMessage = GetDebugMessage(ex); // ex.Message;
                }
                else
                {
                    errorMessage = errorMessage + "\n\r--------------------------------------------\n\r"
                                   + GetDebugMessage(ex); // ex.Message;
                }
                res = ExternalCommandResult.Failed;
            }

            // 对命令的结果进行处理
            switch (res)
            {
                case ExternalCommandResult.Failed:
                    {
                        MessageBox.Show(errorMessage, @"外部命令执行出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // 选择出错的单元格
                        if (errorRange != null)
                        {
                            errorRange.Select();
                        }
                        break;
                    }
                case ExternalCommandResult.Cancelled:
                    {
                        // 由于没有在Excel中没有事务或者回滚，所以直接结束就可以了。
                        break;
                    }
                case ExternalCommandResult.Succeeded:
                    {
                        break;
                    }
            }
        }

        // 参考 Revit AddInManager.AIM 类
        /// <summary> 执行外部命令 </summary>
        /// <param name="addinItem">注意这里的 addinItem 实例是刷新前的程序集中对应的类，
        /// 这里只能用来提取其 FullName 字符串，而不能直接用来执行， 因为虽然它确实可以执行，但是执行的是重新编译前的那个方法。</param>
        /// <param name="assemblyPath"></param>
        /// <param name="excelApp"></param>
        private static ExternalCommandResult RunActiveCommand(IExcelExCommand addinItem, string assemblyPath, Application excelApp)
        {
            string errorMessage = "";
            Range errorRange = null;

            AssemLoader assemLoader = new AssemLoader();
            ExternalCommandResult result;
            try
            {
                assemLoader.HookAssemblyResolve();

                // 重新加载最新的程序集。注意这里只加载了选择了程序集A本身，而没有加载其所引用的程序集。
                // 此 程序集A 所引用的自定义 程序集B 由上面的 HookAssemblyResolve 方法自动来实现；
                // 但是，当所选择的 Execute 方法中没有使用到程序集B 中的方法时，程序集B 就不会被加载；
                // 这就是为什么 程序集A 中所有引用的 zengfy 自定义程序集B，都必须在  Execute() 方法中调用至少一次，以解决在Form.Show()时，出现不能找到或加载前面缺失的程序集B的问题。
                Assembly assembly = assemLoader.LoadAddinsToTempFolder(assemblyPath, false);
                if (null == assembly)
                {
                    MessageBox.Show(@"未能加载程序集", @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    result = ExternalCommandResult.Failed;
                }
                else
                {
                    // !!  注意这里的 addinItem 实例是刷新前的程序集中对应的类，这里只能用来提取其 FullName 字符串，而不能直接用来执行，因为虽然它确实可以执行，但是执行的是重新编译前的那个方法。
                    // !!  这里一定要从最新加载进来的程序集中重新创建对应的外部命令插件
                    IExcelExCommand newExCommand = assembly.CreateInstance(addinItem.GetType().FullName) as IExcelExCommand;

                    if (newExCommand == null)
                    {
                        MessageBox.Show(@"在新加载的程序集中未能找到匹配的方法", @"出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        result = ExternalCommandResult.Failed;
                    }
                    else
                    {
                        result = Execute(newExCommand, excelApp, ref errorMessage, ref errorRange);// 如果在 Execute() 中发现某个程序集不存在，则通过AssemblyResolve 事件手动进行加载
                    }
                }
            }
            catch (Exception ex)
            {
                result = ExternalCommandResult.Failed;
            }
            finally
            {
                assemLoader.UnhookAssemblyResolve();
                assemLoader.CopyGeneratedFilesBack();
            }
            return result;
        }
        
        private static ExternalCommandResult Execute(IExcelExCommand exCommand, Application excelApp, ref string errorMessage,
            ref Range errorRange)
        {
            ExternalCommandResult res = ExternalCommandResult.Failed;
            try
            {
                // 执行操作
                // 如果在执行 Execute()方法时发现某个程序集不存在，则通过AssemblyResolve 事件手动进行加载
                // 所以，所有引用的 zengfy 自定义程序集，都必须在  Execute() 方法中调用至少一次，以解决在Form.Show()时，出现不能找到或加载前面缺失的程序集B的问题。

                // 如果不想通过 AssemblyResolve 来加载缺失的程序集的话，可以在 AddinManager 中自行设计代码，手动在 Execute() 方法之前将要引用的程序集从临时文件夹中通过 Assembly.LoadFile() 进行加载即可。
                res = exCommand.Execute(excelApp: excelApp, errorMessage: ref errorMessage, errorRange: ref errorRange);
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(errorMessage))
                {
                    errorMessage = GetDebugMessage(ex); // ex.Message;
                }
                else
                {
                    errorMessage = errorMessage + "\n\r--------------------------------------------\n\r"
                                   + GetDebugMessage(ex); // ex.Message;
                }
                res = ExternalCommandResult.Failed;
            }

            // 对命令的结果进行处理
            switch (res)
            {
                case ExternalCommandResult.Failed:
                    {
                        MessageBox.Show(errorMessage, @"外部命令执行出错", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        // 选择出错的单元格
                        if (errorRange != null)
                        {
                            errorRange.Select();
                        }
                        break;
                    }
                case ExternalCommandResult.Cancelled:
                    {
                        // 由于没有在Excel中没有事务或者回滚，所以直接结束就可以了。
                        break;
                    }
                case ExternalCommandResult.Succeeded:
                    {
                        break;
                    }
            }
            return res;
        }

        /// <summary> 在调试阶段，为每一种报错显示对应的报错信息及出错位置。 </summary>
        private static string GetDebugMessage(Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(ex.Message);

            // 一直向下提取InnerException
            Exception exInner = ex.InnerException;
            Exception exStack = ex;
            while (exInner != null)
            {
                exStack = exInner;
                sb.AppendLine(exInner.Message);
                exInner = exInner.InnerException;
            }
            // 最底层的出错位置
            sb.AppendLine("\r\n" + exStack.StackTrace);
            return sb.ToString();
        }
    }
}
