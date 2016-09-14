﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using eZvso.AddinManager;
using eZvso.ExternalCommand;

namespace eZvso.AddinManager
{
    internal class AssemblyInfoDllManager
    {

        #region ---   从文件反序列化

        /// <summary> 将 Settings 配置文件中的字符进行反序列化 </summary>
        /// <returns></returns>
        /// <remarks>对于CAD.NET的开发，不要在 IExtensionApplication.Initialize() 方法中执行此操作，否则即使在Initialize时可以正常序列化，
        /// 但是在调用ExternalCommand时还是会出bug，通常的报错为：没有为该对象定义无参数的构造函数。 </remarks>
        public static Dictionary<AddinManagerAssembly, List<MethodInfo>> GetInfosFromFile()
        {
            Dictionary<AddinManagerAssembly, List<MethodInfo>> nodesInfo
                = new Dictionary<AddinManagerAssembly, List<MethodInfo>>(new AssemblyComparer());

            // 提取配置文件中的数据
            AssemblyInfoSettings s = new AssemblyInfoSettings();
            if (!string.IsNullOrEmpty(s.AssemblyInfoSerial))
            {
                // 提取字符
                AssemblyInfos amInfos = StringSerializer.Decode64(s.AssemblyInfoSerial) as AssemblyInfos;

                // 提取数据
                nodesInfo = DeserializeAssemblies(amInfos);
            }

            return nodesInfo;
        }


        private static Dictionary<AddinManagerAssembly, List<MethodInfo>> DeserializeAssemblies(
            AssemblyInfos amInfos)
        {
            Dictionary<AddinManagerAssembly, List<MethodInfo>> nodesInfo;
            nodesInfo = new Dictionary<AddinManagerAssembly, List<MethodInfo>>(new AssemblyComparer());
            //
            if (amInfos != null)
            {
                foreach (string assemblyPath in amInfos.AssemblyPaths)
                {
                    if (File.Exists(assemblyPath))
                    {
                        // 将每一个程序集中的外部命令提取出来
                        List<MethodInfo> m = ExternalCommandHandler.LoadExternalCommandsFromAssembly(assemblyPath);
                        if (m.Any())
                        {
                            Assembly ass = m[0].DeclaringType.Assembly;
                            AddinManagerAssembly amAssembly = new AddinManagerAssembly(assemblyPath, ass);
                            if (nodesInfo.ContainsKey(amAssembly))
                            {
                                nodesInfo[amAssembly] = m;
                            }
                            else
                            {
                                nodesInfo.Add(amAssembly, m);
                            }

                        }
                    }
                }
            }
            return nodesInfo;
        }

        #endregion

        #region ---   序列化到文件

        public static void SaveAssemblyInfosToFile(
            Dictionary<AddinManagerAssembly, List<MethodInfo>> nodesInfo)
        {
            // 转换为可序列化的数据
            List<string> assemblyPaths = nodesInfo.Select(r => r.Key.Path).ToList();
            AssemblyInfos amInfos = new AssemblyInfos() { AssemblyPaths = assemblyPaths.ToArray() };

            // 序列化
            string amInfosString = StringSerializer.Encode64(amInfos);

            // 保存到物理存储中
            AssemblyInfoSettings s = new AssemblyInfoSettings();
            s.AssemblyInfoSerial = amInfosString;
            s.Save();
        }


        #endregion

    }
}
