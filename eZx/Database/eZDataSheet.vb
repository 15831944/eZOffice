Imports Microsoft.Office.Interop.Excel
Imports eZstd.eZexcelAPI
Imports ExcelAddIn_zfy.DataSheet
Imports System.ComponentModel

Public Class eZDataSheet

#Region "  ---  Declarations & Definitions"

#Region "  ---  Types"

#End Region

#Region "  ---  Events"

#End Region

#Region "  ---  Constants"

#End Region

#Region "  ---  Properties"

    ''' <summary>
    ''' 代表此数据库的工作表对象
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkSheet As Excel.Worksheet

    ''' <summary>
    ''' 此工作表是否符合数据库的格式规范
    ''' </summary>
    Private F_IsFormated As Boolean

    Public Property Fields As BindingList(Of DataField)

    '''<summary> 此字段名称本身的数据类型。
    ''' 一般情况下，一个字段的名称只要是一个字符就可以了，但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </summary>
    Public Property FieldType As eZDataType
#End Region

#Region "  ---  Fields"

#End Region

#End Region

#Region "  ---  构造函数与窗体的加载、打开与关闭"

    ''' <summary>
    ''' 构造函数
    ''' </summary>
    ''' <param name="WkSheet"></param>
    ''' <param name="List_FieldInfo"></param>
    ''' <param name="FieldType">字段名称本身的数据类型, 一般情况下，一个字段的名称只要是一个字符就可以了，
    ''' 但是它也可以代表具体的含义，比如在具体某一天的日期“2016/2/6” </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal WkSheet As Worksheet, ByVal List_FieldInfo As BindingList(Of DataField), _
                   Optional FieldType As eZDataType = eZDataType.字符)
        Me.WorkSheet = WkSheet
        Me.Fields = List_FieldInfo
        Me.FieldType = FieldType
    End Sub
#End Region

End Class
