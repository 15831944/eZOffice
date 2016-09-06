namespace eZx.Database
{
    /// <summary> Excel中可以辨别的数据类型，用来进行字段名或者某字段下的数据的类别判断 </summary>
    public enum eZDataType
    {
        字符, // String
        日期, // DateTime
        整数, // Int64
        浮点数 // Double
    }
}