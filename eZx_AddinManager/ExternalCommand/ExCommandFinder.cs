using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using eZx.AddinManager;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace eZx.AddinManager
{
    /// <summary> 将指定程序集中的 IExternalCommand 类提取出来 </summary>
    public static class ExCommandFinder
    {
        /// <summary> 将程序集文件加载到内存，并且提取出其中的 CAD 外部命令 </summary>
        /// <param name="assemblyPath"></param>
        /// <returns></returns>
        public static List<IExternalCommand> RetriveExternalCommandsFromAssembly(string assemblyPath)
        {
            //先将插件拷贝到内存缓冲。一般情况下，当加载的文件大小大于2^32 byte (即4.2 GB），就会出现OutOfMemoryException，在实际测试中的极限值为630MB。
            byte[] buff = File.ReadAllBytes(assemblyPath);

            //不能直接通过LoadFrom或者LoadFile，而必须先将插件拷贝到内存，然后再从内存中Load
            Assembly asm = Assembly.Load(buff);

            //
            // loadReferences(asm, assemblyPath);
            return GetExternalCommandClass(asm);
        }

        /// <summary> 添加程序集的引用项 </summary>
        /// <param name="asm"></param>
        /// <param name="assemblyPath"></param>
        private static void loadReferences(Assembly asm, string assemblyPath)
        {
            // 提取当前文件夹中所有后缀为 .dll 与 .exe 的文件
            var dllFiles = new FileInfo(assemblyPath).Directory
                .GetFiles("*.*", SearchOption.TopDirectoryOnly)
                .Where(file => string.Compare(file.Extension, ".dll", StringComparison.OrdinalIgnoreCase) == 0
                || string.Compare(file.Extension, ".exe", StringComparison.OrdinalIgnoreCase) == 0).Where(
                file => !file.Name.StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase)
                && !file.Name.StartsWith("System", StringComparison.OrdinalIgnoreCase)).ToArray();

            // 排除以 System 与 Microsoft 开头的程序集
            var references = asm.GetReferencedAssemblies()
                .Where(ass => !ass.Name.StartsWith("Microsoft", StringComparison.OrdinalIgnoreCase)
                && !ass.Name.StartsWith("System", StringComparison.OrdinalIgnoreCase)).ToArray();

            // 确定哪些引用的程序集文件需要进行加载
            var fileNames = dllFiles.Select(r => r.Name.Substring(0, r.Name.Length - r.Extension.Length)).ToList();

            foreach (var asmName in references)
            {
                int index;
                if ((index = fileNames.IndexOf(asmName.Name)) >= 0)
                {
                    var result = Assembly.LoadFile(dllFiles[index].FullName);

                    // var refAss = Assembly.Load(asmName);
                    // var refAss = Assembly.Load(dllFiles[index].FullName);
                }
            }
        }

        private static List<IExternalCommand> GetExternalCommandClass(Assembly ass)
        {
            List<IExternalCommand> ecClasses = new List<IExternalCommand>();
            var classes = ass.GetTypes();
            foreach (Type cls in classes)
            {
                if (cls.GetInterfaces().Any(r => r == typeof(IExternalCommand))) // 说明这个类实现了 CAD 的命令接口
                {
                    // 寻找此类中所实现的那个 Execute 方法
                    Type[] paraTypes = new Type[3]
                    {typeof (Application), typeof (string).MakeByRefType(), typeof (Range).MakeByRefType()};
                    //
                    MethodInfo m = cls.GetMethod("Execute", paraTypes);
                    //
                    if (m != null && m.IsPublic)
                    {
                        // 生成一个实例并转换为接口
                        var ins = ass.CreateInstance(cls.FullName);
                        IExternalCommand exC = ins as IExternalCommand;

                        if (exC != null)
                        {
                            ecClasses.Add(exC);
                        }
                    }
                }
            }
            return ecClasses;
        }
    }
}