using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Visio;
using Application = Microsoft.Office.Interop.Visio.Application;

namespace eZvso.CurveMaker
{
    internal class Drawer
    {
        /// <summary>
        /// 绘制一般的样条曲线
        /// </summary>
        /// <param name="vsoPage"></param>
        /// <param name="points"></param>
        /// <param name="tolerance">How closely the path of the new shape must approximate the given points.
        /// The error from the points to the path of the resulting shape is roughly within tolerance. 
        /// When the number of points is large, the actual error may sometimes exceed the prescribed tolerance.</param>
        public static void DrawSplineOnPage(Page vsoPage, List<LocationPoint> points, double tolerance)
        {
            double[] controllingPoints = LocationPoint.ToSafeArray(points);
            //
            Shape vsoShape = vsoPage.DrawSpline(
                xyArray: controllingPoints,
                Tolerance: tolerance,
                Flags: (short)VisDrawSplineFlags.visSplinePeriodic);
        }



        /// <summary>
        /// 绘制多段线折线
        /// </summary>
        /// <param name="vsoPage"></param>
        /// <param name="points"></param>
        public static void DrawPolylineOnPage(Page vsoPage, List<LocationPoint> points)
        {
            double[] controllingPoints = LocationPoint.ToSafeArray(points);
            //
            Shape vsoShape = vsoPage.DrawPolyline(controllingPoints, (short)VisDrawSplineFlags.visPolyline1D);
        }

        /// <summary> 绘制一条 贝塞尔曲线  </summary>
        public static void DrawBezierOnPage(Page vsoPage, List<LocationPoint> points, int degree)
        {
            double segments = (double)(points.Count - 1) / degree;
            // 生成的最终曲线中有 n 段贝塞尔曲线
            int ns = (int)segments;
            if (Math.Abs(segments - ns) > 0.0000001)
            {
                MessageBox.Show(@"用来绘制贝塞尔曲线的控制点个数 np 必须满足 np = ns * degree + 1，其中 ns 为一个正整数，" +
                                @"表示生成的最终曲线中的贝塞尔曲线段的段数");
                return;
            }

            // 如果输入参数符合要求，则开始绘图
            DrawBezier_Example(vsoPage.Application);
        }

        #region ---   测试
        /// <summary> 绘制一个 NURBS 曲线 </summary>
        /// <param name="vsoApp"></param>
        private static void FitSmoothTest(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            if (doc != null)
            {
                Selection s = vsoApp.Windows[1].Selection;
                Shape shp = s[1];
                shp.FitCurve(0.5, (short)VisDrawSplineFlags.visSplinePeriodic);
            }
        }

        /// <summary> 绘制一个 NURBS 曲线 </summary>
        private static void DrawNURBSTest(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            if (doc != null)
            {
                int DiagramServices = doc.DiagramServicesEnabled;
                doc.DiagramServicesEnabled = (int)VisDiagramServices.visServiceVersion140 +
                                             (int)VisDiagramServices.visServiceVersion150;
                // 8 个控制点
                Double[] ControlPoints1 = new double[16]
                {
            0.590551  ,5.41339   ,0.536328  ,5.86033   ,
            1.39326   ,6.75451   ,2.27205   ,6.10485   ,
            2.93392   ,5.47316   ,3.91569   ,5.26246   ,
            4.26121   ,5.63617   ,4.33071   ,5.70866   ,
                };
                // 
                Double[] Knots1 = new double[9]
                {
            0      , 0      , 0      ,
            0.5      , 1.68063, 2.18785,
            3.35816, 4.49301, 4.8251 ,
                };

                Page pg = doc.Pages[1];
                Shape shp = pg.DrawNURBS(degree: 3,
                     Flags: (short)VisDrawSplineFlags.visSplinePeriodic,
                     xyArray: ControlPoints1,
                     knots: Knots1);
                // Restore diagram services
                doc.DiagramServicesEnabled = DiagramServices;
            }
        }

        /// <summary> 绘制一个 Spline </summary>
        private static void DrawSplineTest(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            Shape vsoShape;
            int n = 20;
            double[] xyArray = new double[4] { 1, 2, 3, 4 }; //表示两个点(1,2)与(3,4)
            double[] adblXYPoints = new double[n * 2];
            //
            for (int intCounter = 1; intCounter <= n; intCounter++)
            {
                // Set x components (array elements 1,3,5,7,9) to 1,2,3,4,5 
                adblXYPoints[(intCounter * 2) - 2] = (double)intCounter / 2 * Math.PI;

                //Set y components (array elements 2,4,6,8,10) to f(i) 
                // adblXYPoints[intCounter * 2 - 1] = (intCounter * intCounter) - (7 * intCounter) + 15;
                adblXYPoints[intCounter * 2 - 1] = Math.Sin((double)intCounter / 2 * Math.PI); // (intCounter * intCounter) - (7 * intCounter) + 15;
            }

            //
            Page pg = doc.Pages[1];
            vsoShape = pg.DrawSpline(
                xyArray: adblXYPoints,
                Tolerance: 0.25,
                Flags: (short)VisDrawSplineFlags.visPolyarcs);
        }

        /// <summary> 绘制一条 贝塞尔曲线  </summary>
        private static void DrawBezier_Example(Application vsoApp)
        {
            Document doc = vsoApp.ActiveDocument;
            Shape vsoShape;
            int n = 5;
            double[] adblXYPoints = new double[n * 2];
            //
            for (int intCounter = 1; intCounter <= n; intCounter++)
            {
                // Set x components (array elements 1,3,5,7,9) to 1,2,3,4,5 
                adblXYPoints[(intCounter * 2) - 2] = (double)intCounter;
                //Set y components (array elements 2,4,6,8,10) to f(i) 
                adblXYPoints[intCounter * 2 - 1] = (intCounter * intCounter) - (7 * intCounter) + 15;
            }

            //
            Page pg = doc.Pages[1];

            vsoShape = pg.DrawBezier(
                xyArray: adblXYPoints,
                degree: 3,
                // The Flags argument is a bitmask that specifies options for drawing the new shape. Its value should be zero (0) or visSpline1D (8).
                Flags: (short)VisDrawSplineFlags.visSpline1D);
        }

        #endregion
    }
}
