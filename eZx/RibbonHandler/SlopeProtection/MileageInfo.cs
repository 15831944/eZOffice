using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using eZstd.Miscellaneous;

namespace eZx.RibbonHandler.SlopeProtection
{

    public class MileageInfo
    {
        public float Mileage { get; private set; }
        public MileageInfoType Type { get; private set; }
        public double SpLength { get; set; }

        public static Dictionary<string, MileageInfoType> TypeMapping = new Dictionary<string, MileageInfoType>
        {
            {"定位", MileageInfoType.Located},
            {"测量", MileageInfoType.Measured},
            {"插值", MileageInfoType.Interpolated},
        };

        public MileageInfo(float mileage, MileageInfoType type, double slopeLength)
        {
            Mileage = mileage;
            Type = type;
            SpLength = slopeLength;
        }

        /// <summary> 用新的数据替换对象中的原数据 </summary>
        public void Override(MileageInfo newSection)
        {
            Mileage = newSection.Mileage;
            Type = newSection.Type;
            SpLength = newSection.SpLength;
        }

        /// <summary>
        /// 将 边坡横断面集合转换为二维数组，以用来写入 Excel
        /// </summary>
        /// <param name="slopes"></param>
        /// <returns></returns>
        public static object[,] ConvertToArr(IList<MileageInfo> slopes)
        {
            var res = new object[slopes.Count(), 3];
            var keys = TypeMapping.Keys.ToArray();
            var values = TypeMapping.Values.ToArray();

            var r = 0;
            foreach (var slp in slopes)
            {
                res[r, 0] = slp.Mileage;
                res[r, 1] = keys[Array.IndexOf(values, slp.Type)];
                res[r, 2] = slp.SpLength;
                r += 1;
            }
            return res;
        }
    }


    /// <summary> 里程桩号小的在前面 </summary>
    public class MileageCompare : IComparer<MileageInfo>
    {
        public int Compare(MileageInfo x, MileageInfo y)
        {
            return x.Mileage.CompareTo(y.Mileage);
        }
    }

    public enum MileageInfoType
    {
        Located,
        Measured,
        Interpolated,
    }
}
