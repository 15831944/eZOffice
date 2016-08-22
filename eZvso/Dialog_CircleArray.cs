// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Xml.Linq;
using Visio = Microsoft.Office.Interop.Visio;
using System.Text;
using System.Linq;
// End of VB project level imports

using System.Windows.Forms;

namespace eZvso
{
    public partial class Dialog_CircleArray
    {
        public Dialog_CircleArray()
        {
            InitializeComponent();
        }
        private UInt16 Num = 4;
        private double Angle = 360;
        private bool blnCenter = true;
        private bool blnPreserveDirection = false;


        public new DialogResult ShowDialog(UInt16 Num = 4, double Angle = 360, bool blnCenter = true, bool blnPreserveDirection = true)
        {
            DialogResult res = base.ShowDialog();
            if (res == DialogResult.OK)
            {
                Dialog_CircleArray with_1 = this;
                Num = with_1.Num;
                Angle = with_1.Angle;
                blnCenter = with_1.blnCenter;
                blnPreserveDirection = with_1.blnPreserveDirection;
            }
            return res;

        }

        #region    ---   事件处理
        public void txtAngle_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            double.TryParse(txt.Text, out Angle);
        }
        public void txtNum_TextChanged(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;
            UInt16.TryParse(txt.Text, out Num);
        }

        public void CheckBox_preserveDirection_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox box = (CheckBox)sender;
            if (box.Checked)
            {
                this.blnPreserveDirection = true;
            }
            else
            {
                this.blnPreserveDirection = false;
            }
        }
        public void RadioButton_Center_CheckedChanged(object sender, EventArgs e)
        {
            if (this.RadioButton_Center.Checked)
            {
                this.blnCenter = true;
            }
            else
            {
                this.blnCenter = false;
            }

        }

        public void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion

    }
}
