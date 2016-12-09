using System.Configuration;
using Office = Microsoft.Office.Core;

//Application settings wrapper class. This class defines the settings we intend to use in our application.

namespace eZx
{
    public sealed class HelpLocationSettings : ApplicationSettingsBase
    {
        /// <summary> Office 开发帮助文档的文件夹的绝对路径 </summary>
        [UserScopedSetting(),DefaultSettingValue(
             "F:\\Software\\Programming\\VB.NET\\VB.NET与二次开发\\Visual Basic.NET与Office开发\\开发工具与资料\\Office 2013 VBA Documentation"
             )]
        public string OfficeHelp
        {
            get { return this["OfficeHelp"].ToString(); }
            set { this["OfficeHelp"] = value; }
        }

        /// <summary> Excel 开发帮助文档的绝对路径 </summary>
        [UserScopedSetting(),DefaultSettingValue(
             "F:\\Software\\Programming\\VB.NET\\VB.NET与二次开发\\Visual Basic.NET与Office开发\\开发工具与资料\\Office 2013 VBA Documentation\\Excel 2013 Developer Documentation.chm"
             )]
        public string ExcelHelp
        {
            get { return this["ExcelHelp"].ToString(); }
            set { this["ExcelHelp"] = value; }
        }
    }
}