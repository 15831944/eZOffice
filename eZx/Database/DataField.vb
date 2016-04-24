Public Module DataSheet
    ''' <summary>
    ''' 数据表中每一个字段的信息
    ''' </summary>
    Public Class DataField

        ''' <summary> 此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        ''' 在Excel 2010中，最大的列号为16384=2^14。 </summary>
        Public Property ColumnIndex As UInt16

        ''' <summary> 字段名称 </summary>
        Public Property Name As String

        ''' <summary> 此列数据的类型 </summary>
        Public Property DataType As eZDataType

        ''' <summary> 是否允许空值，如果为False，则会自动将其设置为其默认值 </summary>
        Public Property NullAllowed As Boolean

        '''<summary>构造函数</summary>
        ''' <param name="name">字段名称</param>
        ''' <param name="ColumnIndex">此字段在数据库中的列号下标，比如第一列(A列)的数据的ColumnIndex为1。
        ''' 在Excel 2010中，最大的列号为16384=2^14。 </param>
        ''' <param name="dataType">此列的数据类型</param>
        ''' <param name="nullAllowed">是否允许有空值</param>
        ''' <remarks></remarks>
        Public Sub New(name As String, ByVal ColumnIndex As UInt16, Optional dataType As eZDataType = Nothing, _
                       Optional ByVal nullAllowed As Boolean = True)
            With Me
                .Name = name
                .ColumnIndex = ColumnIndex
                .DataType = dataType
                .NullAllowed = nullAllowed
                'If fieldtype = Nothing Then
                '    .FieldType = eZDataType.字符
                'End If
                If dataType = Nothing Then
                    .DataType = eZDataType.字符
                End If
            End With
        End Sub

    End Class

    Public Enum eZDataType
        字符      ' String
        日期      ' DateTime
        整数      ' Int64
        浮点数    ' Double 
    End Enum

    Public Function IsCompatible(ByVal CheckedData As String, ByVal ezType As eZDataType) As Boolean
        Dim blnIsCompatible As Boolean = True
        Select Case ezType
            Case eZDataType.整数
                Dim v As Int64
                blnIsCompatible = Int64.TryParse(CheckedData, v)
            Case eZDataType.浮点数
                Dim v As Double
                blnIsCompatible = Double.TryParse(CheckedData, v)
            Case eZDataType.日期
                Dim v As Date
                blnIsCompatible = Date.TryParse(CheckedData, v)
        End Select
        Return blnIsCompatible
    End Function

End Module
