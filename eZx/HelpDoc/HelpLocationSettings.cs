using System.Configuration;
using Office = Microsoft.Office.Core;

//Application settings wrapper class. This class defines the settings we intend to use in our application.

namespace eZx
{
    public sealed class HelpLocationSettings : ApplicationSettingsBase
    {
        [UserScopedSetting(),
         DefaultSettingValue(
             "F:\\Software\\Programming\\VB.NET\\VB.NET与二次开发\\Visual Basic.NET与Office开发\\开发工具与资料\\Office 2013 VBA Documentation"
             )]
        public string OfficeHelp
        {
            get { return this["OfficeHelp"].ToString(); }
            set { this["OfficeHelp"] = value; }
        }

        [UserScopedSetting(),
         DefaultSettingValue(
             "F:\\Software\\Programming\\VB.NET\\VB.NET与二次开发\\Visual Basic.NET与Office开发\\开发工具与资料\\Office 2013 VBA Documentation\\Excel 2013 Developer Documentation.chm"
             )]
        public string ExcelHelp
        {
            get { return this["ExcelHelp"].ToString(); }
            set { this["ExcelHelp"] = value; }
        }
    }
}