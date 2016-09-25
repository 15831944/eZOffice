using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace eZx.ExternalCommand
{
    public enum ExternalCommandResult
    {
        Cancelled = 0,
        Succeeded = 1,
        Failed = 2,
    }

    /// <summary> 用来进行AddinManager快速调试的接口。实现此接口的类必须有一个无参数的构造函数 </summary>
    public interface IExternalCommand
    {
        /// <summary> Excel AddinManger 快速调试插件 </summary>
        /// <param name="excelApp">Excel当前程序</param>
        /// <param name="errorMessage">当返回值为<see cref="ExternalCommandResult.Failed"/>时，这个属性代表给出的报错信息。</param>
        /// <param name="errorRange">当返回值为<see cref="ExternalCommandResult.Failed"/>时，这个属性代表与出错内容相关的任何对象。</param>
        /// <returns></returns>
        ExternalCommandResult Execute(Application excelApp, ref string errorMessage, ref Range errorRange);
    }
}
