using System;
using System.Windows.Forms;
using eZvso.CurveMaker;
using Microsoft.Office.Interop.Visio;
using Microsoft.Office.Tools.Ribbon;
using Application = Microsoft.Office.Interop.Visio.Application;
using Office = Microsoft.Office.Core;

namespace eZvso
{
    public partial class Ribbon_eZvso
    {
        #region   ---  Fields

        private Application _app;

        /// <summary>
        /// 当前正在进行编辑的Master对象（不是指Master所对应的实例形状）
        /// </summary>
        /// <remarks></remarks>
        private Master MasterInEdit;

        /// <summary>
        ///  旋转中心点相对于实例形状范围界定框的左下角点的X位置
        /// </summary>
        private double MasterBase_LocPinX = 0.5;

        /// <summary>
        /// 旋转中心点相对于实例形状范围界定框的左下角点的Y位置
        /// </summary>
        private double MasterBase_LocPinY = 0.5;

        /// <summary>
        /// 进行放置阵列的对话框
        /// </summary>
        /// <remarks></remarks>
        Dialog_CircleArray Dlg_CircleArray;

        #endregion

        #region   ---  构造函数与窗体的加载、打开与关闭

        private void Ribbon_eZvso_Load(Object sender, RibbonUIEventArgs e)
        {
            _app = Globals.ThisAddIn.Application;
            _app.WindowActivated += App_WindowActivated;
        }

        #endregion

        #region    ---   绘图

        /// <summary>
        /// 将形状原位粘贴到组
        /// </summary>
        public void AddToGroup(object sender, RibbonControlEventArgs e)
        {
            _app.ActiveWindow.Selection.AddToGroup();
        }

        /// <summary>
        /// 图形的平移、矩形阵列、面积与周长
        /// </summary>
        public void btnMove_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton btn = (RibbonButton)sender;
            _app.Addons[short.Parse(Convert.ToString(btn.Tag))].Run(""); //5 表示平移，6表示形状的面积与周长，7表示阵列
            //如果要执行阵列命令，也可以用：    App.DoCmd(1354)  ' VisUICmds 常量中的 visCmdToolsArrayShapesAddOn  命令
        }

        /// <summary>
        /// 旋转阵列
        /// </summary>
        public void CircleArray(object sender, RibbonControlEventArgs e)
        {
            Selection sel = _app.ActiveWindow.Selection;
            if (sel.Count > 0)
            {
                if (Dlg_CircleArray == null)
                {
                    Dlg_CircleArray = new Dialog_CircleArray();
                }
                double angle = 0; // 旋转阵列的总角度
                UInt16 n = 0; // 旋转阵列的个数
                bool blnPreserveDirection = false; //是否保留对象的角度方向
                //
                //DialogResult res = Dlg_CircleArray.ShowDialog(Num: ref n, Angle: ref angle, blnPreserveDirection: ref blnPreserveDirection);
                DialogResult res = Dlg_CircleArray.ShowDialog(Num: n, Angle: angle,
                    blnPreserveDirection: blnPreserveDirection);

                if (res == DialogResult.OK)
                {
                    Shape shp;
                    if (sel.Count == 1)
                    {
                        shp = sel[1];
                    }
                    else
                    {
                        shp = sel.Group();
                    }
                    double baseX = Convert.ToDouble(shp.Cells["PinX"].ResultIU); // 图形的旋转中心在页面中的绝对X坐标
                    double baseY = Convert.ToDouble(shp.Cells["PinY"].ResultIU); // 图形的旋转中心在页面中的绝对Y坐标
                    double OrigionalAngle = Convert.ToDouble(shp.Cells["Angle"].Result[VisUnitCodes.visDegrees]);
                    double Width = Convert.ToDouble(shp.Cells["Width"].ResultIU);
                    double Height = Convert.ToDouble(shp.Cells["Height"].ResultIU);
                    string strLocPinX = Convert.ToString(shp.Cells["LocPinX"].Formula);
                    string strLocPinY = Convert.ToString(shp.Cells["LocPinY"].Formula);
                    double WidthScale = 0;
                    double HeightScale = 0;

                    try
                    {
                        WidthScale = double.Parse(strLocPinX.Substring(6)); // 图形的旋转中心相对于图形的左下角点的位置： Width*0.5
                        HeightScale = double.Parse(strLocPinY.Substring(7)); // 图形的旋转中心相对于图形的左下角点的位置： Height*0.5
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("请在ShapeSheet中以相对值的形式来表达 LocPinX 与 LocPinY 的值", "Error", MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }
                    double OriginalCenterX = baseX - Width * WidthScale + Width * 0.5; // 图形的中心点在页面中的绝对X坐标
                    double OriginalCenterY = baseY - Height * HeightScale + Height * 0.5; // 图形的中心点在页面中的绝对Y坐标
                    double r = Math.Sqrt(Math.Pow(baseX - OriginalCenterX, 2) + Math.Pow(baseY - OriginalCenterY, 2));
                    // ------------------------ 开始复制形状  ----------------------
                    Shape NewShape = default(Shape);
                    _app.ShowChanges = false;
                    for (UInt16 i = 1; i <= n - 1; i++)
                    {
                        double deltaA = Convert.ToDouble(angle / n * i / 180 * Math.PI); // 单位为弧度
                        NewShape = shp.Duplicate();
                        //将形状移动回原位
                        NewShape.Cells["PinX"].ResultIU = baseX;
                        NewShape.Cells["PinY"].ResultIU = baseY;
                        if (blnPreserveDirection) //是否保留对象的角度方向
                        {
                            if (r > 0) // 此时新图形与原图形在同一个位置，不用作任何的移动，而且下面的alpha角算出来为无穷，因为分母r为0.
                            {
                                double alpha = Math.Asin((OriginalCenterY - baseY) / r);
                                double NewCenterX = baseX + r * Math.Cos(deltaA + alpha); // 注意三角函数计算时的单位为弧度
                                double NewCenterY = baseY + r * Math.Sin(deltaA + alpha);
                                //新形状的中心点在页面中的绝对坐标值
                                NewShape.Cells["PinX"].ResultIU = NewCenterX - OriginalCenterX + baseX;
                                NewShape.Cells["PinY"].ResultIU = NewCenterY - OriginalCenterY + baseY;
                            }
                        }
                        else
                        {
                            NewShape.Cells["Angle"].Result[VisUnitCodes.visDegrees] = OrigionalAngle +
                                                                                      deltaA / Math.PI * 180;
                        }
                    }

                    _app.ShowChanges = true;
                }
            }
            else
            {
                MessageBox.Show(@"请先选择至少一个图形对象。");
            }
        }

        #endregion

        #region    ---   主控形状编辑

        /// <summary>
        /// 绘制主控形状的基点位置
        /// </summary>
        public void btnMasterBase_Click(object sender, RibbonControlEventArgs e)
        {
            Shape shape = MasterInEdit.DrawLine(0, 0, 0, 0);
            // MasterInEdit.
        }

        public void btnLocPin_Click(object sender, RibbonControlEventArgs e)
        {
        }

        #endregion

        #region    ---   事件处理

        private void App_WindowActivated(Window Window)
        {
            if (Window.SubType == 64) // 主控形状绘图页窗口。（通过“文档模具”中右键，“编辑主控形状”所进入的窗口）
            {
                this.btnMasterBase.Enabled = true;
                MasterInEdit = Window.Master as Master;
            }
            else
            {
                this.btnMasterBase.Enabled = false;
                MasterInEdit = null;
            }
        }

        #endregion

        private void button_FunctionCurve_Click(object sender, RibbonControlEventArgs e)
        {
            frm_CurveParameter f =  frm_CurveParameter.GetUniqueInstance(Globals.ThisAddIn.Application);
            f.ShowDialog();
        }
    }
}