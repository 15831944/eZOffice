using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;

namespace eZvso.eZvso_API
{
    public static class ShapeSearching
    {
        /// <summary>
        /// 按逐级深入的方式获得形状的所有子形状
        /// </summary>
        /// <param name="shp"></param>
        /// <param name="includeMyself">返回的所有子形状的集合中是否包含自身</param>
        /// <returns></returns>
        public static List<Shape> GetAllShapes(Shape shp, bool includeMyself = true)
        {
            List<Shape> allNestedShapes = new List<Shape>();
            List<Shape> childShapes = new List<Shape>(); // 每一层搜索的下一级子形状
            if (includeMyself)
            {
                allNestedShapes.Add(shp);
            }

            // 第一级
            foreach (Shape childShp in shp.Shapes)
            {
                childShapes.Add(childShp);
            }
            allNestedShapes.AddRange(childShapes);

            // 看是否需要进行更深层的搜索
            while (childShapes.Count > 0)
            {
                childShapes = AddNextChildShapes(allNestedShapes, childShapes);
            }
            return allNestedShapes;
        }

        /// <summary>
        /// 搜索某一级形状集合中的下一级的所有子形状
        /// </summary>
        /// <param name="totalShapes"></param>
        /// <param name="shps">要搜索子形状的某一级形状的集合</param>
        /// <returns>搜索到的下一级的所有子形状</returns>
        private static List<Shape> AddNextChildShapes(List<Shape> totalShapes, List<Shape> shps)
        {
            List<Shape> childShapes = new List<Shape>();
            foreach (Shape shp in shps)
            {
                foreach (Shape childShp in shp.Shapes)
                {
                    childShapes.Add(childShp);
                }
            }
            totalShapes.AddRange(childShapes);
            return childShapes;
        }
    }
}
