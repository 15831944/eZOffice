using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eZx.Debug.工程量清单
{
    public enum Matched
    {
        匹配,
        未匹配,
        未包含
    }
    public class Item
    {

        public string 子目号 { get; set; }
        public string 子目名称 { get; set; }
        public string 项目特征 { get; set; }
        public string 单位 { get; set; }
        public Matched 备注 { get; set; }
        public Item(string 子目号, string 子目名称, string 项目特征, string 单位)
        {
            this.子目号 = 子目号;
            this.子目名称 = 子目名称;
            this.项目特征 = 项目特征;
            this.单位 = 单位;
        }

        public override string ToString()
        {
            return 子目号;
        }

        public object[] ToArray()
        {
            return new object[] { Enum.GetName(typeof(Matched), 备注), 子目号, 子目名称, 项目特征, 单位 };
        }
    }
}
