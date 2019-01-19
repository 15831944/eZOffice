using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Core;

namespace eZx.CppDoc
{
    public enum DocMemberType
    {
        Function,
        Property,
    }

    /// <summary> 一个 </summary>
    public class CppDocMember
    {
        /// <summary> 成员名称 </summary>
        public string MemberName { get; }

        /// <summary> 成员名称 </summary>
        public DocMemberType MemberType { get; }

        /// <summary> 返回值类型 </summary>
        public string ReturnType { get; }
        /// <summary> 成员签名 </summary>
        public string MemberSignature { get; }
        /// <summary> 成员描述 </summary>
        public string Description { get; }

        public CppDocMember(string memberName, DocMemberType memberType = DocMemberType.Function,
            string memberSignature = null,
            string returnType = "void",
            string description = null)
        {
            MemberName = memberName;
            MemberType = memberType;
            MemberSignature = memberSignature;
            ReturnType = returnType;
            Description = description;
        }

        private static string pattern = @"(.+)\((.*)\)";
        public static string ExtractMemberNameFromSignature(string signature)
        {
            Regex reg = new Regex(pattern);
            Match m = reg.Match(signature);
            if (m.Success)
            {
                //string inputPara = m.Groups[2].Value;
                return m.Groups[1].Value;
            }
            return null;
        }

        public string GetMemberTypeName()
        {
            return Enum.GetName(typeof(DocMemberType), MemberType);
        }

        public override string ToString()
        {
            return $"{MemberName}: {GetMemberTypeName()}";
        }

    }
}
