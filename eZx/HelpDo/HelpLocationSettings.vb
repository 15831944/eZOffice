'Application settings wrapper class. This class defines the settings we intend to use in our application.
Imports System.Configuration

NotInheritable Class HelpLocationSettings
    Inherits ApplicationSettingsBase

    <UserScopedSettingAttribute(), 
        DefaultSettingValueAttribute("F:\Software\Programming\VB.NET\VB.NET与二次开发\Visual Basic.NET与Office开发\开发工具与资料\Office 2013 VBA Documentation")>
    Public Property OfficeHelp() As String
        Get
            Return Me("OfficeHelp")
        End Get
        Set(ByVal value As String)
            Me("OfficeHelp") = value
        End Set
    End Property

    <UserScopedSettingAttribute(), 
        DefaultSettingValueAttribute("F:\Software\Programming\VB.NET\VB.NET与二次开发\Visual Basic.NET与Office开发\开发工具与资料\Office 2013 VBA Documentation\Excel 2013 Developer Documentation.chm")>
    Public Property ExcelHelp() As String
        Get
            Return Me("ExcelHelp")
        End Get
        Set(ByVal value As String)
            Me("ExcelHelp") = value
        End Set
    End Property

End Class
