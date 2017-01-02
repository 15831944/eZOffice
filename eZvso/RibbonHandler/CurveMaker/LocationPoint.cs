using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eZvso.RibbonHandler.CurveMaker
{
    /// <summary>
    /// 用来绘制曲线的坐标点，其单位为Visio的内部单位英寸 1 inch = 25.4 mm
    /// </summary>
    internal class LocationPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public override string ToString()
        {
            return $"({X} ,{Y})";
        }

        /// <summary>
        /// 将坐标点集合转换为一维数组，以用作 Draw方法的 xyArray 参数
        /// </summary>
        /// <returns>数组中每两个连续的元素组成一个点的x、y坐标值</returns>
        /// <remarks>and weights arrays should be of type SAFEARRAY of 8-byte floating point values passed by  reference (VT_R8|VT_ARRAY|VT_BYREF).
        ///  This is how Microsoft Visual Basic passes arrays to Automation objects.</remarks>
        public static double[] ToSafeArray(IList<LocationPoint> points)
        {
            double[] pointsArray = new double[points.Count * 2];
            for (int i = 0; i < points.Count; i++)
            {
                pointsArray[2 * i] = points[i].X;
                pointsArray[2 * i + 1] = points[i].Y;
            }
            return pointsArray;
        }
    }
}
