using System;
using Microsoft.Office.Interop.Excel;

namespace eZx_API.UserControls
{
    partial class CurveRangeLocator
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button_srcXY = new System.Windows.Forms.Button();
            this.rangeGetorX = new eZx_API.UserControls.RangeGetor();
            this.rangeGetorY = new eZx_API.UserControls.RangeGetor();
            this.SuspendLayout();
            // 
            // button_srcXY
            // 
            this.button_srcXY.Location = new System.Drawing.Point(216, 3);
            this.button_srcXY.Margin = new System.Windows.Forms.Padding(2);
            this.button_srcXY.Name = "button_srcXY";
            this.button_srcXY.Size = new System.Drawing.Size(17, 45);
            this.button_srcXY.TabIndex = 15;
            this.button_srcXY.Text = "XY";
            this.button_srcXY.UseVisualStyleBackColor = true;
            this.button_srcXY.Click += new System.EventHandler(this.button_srcXY_Click);
            // 
            // rangeGetorX
            // 
            this.rangeGetorX.ButtonText = "X";
            this.rangeGetorX.LabelText = "X数据源";
            this.rangeGetorX.Location = new System.Drawing.Point(4, 3);
            this.rangeGetorX.Name = "rangeGetorX";
            this.rangeGetorX.Size = new System.Drawing.Size(212, 22);
            this.rangeGetorX.TabIndex = 20;
            this.rangeGetorX.RangeChanged += new System.Action<object, Microsoft.Office.Interop.Excel.Range>(this.rangeGetorX_RangeChanged);
            // 
            // rangeGetorY
            // 
            this.rangeGetorY.ButtonText = "Y";
            this.rangeGetorY.LabelText = "Y数据源";
            this.rangeGetorY.Location = new System.Drawing.Point(4, 28);
            this.rangeGetorY.Name = "rangeGetorY";
            this.rangeGetorY.Size = new System.Drawing.Size(212, 22);
            this.rangeGetorY.TabIndex = 18;
            this.rangeGetorY.RangeChanged += new System.Action<object, Microsoft.Office.Interop.Excel.Range>(this.rangeGetorY_RangeChanged);
            // 
            // CurveRangeLocator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.rangeGetorX);
            this.Controls.Add(this.rangeGetorY);
            this.Controls.Add(this.button_srcXY);
            this.MaximumSize = new System.Drawing.Size(235, 50);
            this.MinimumSize = new System.Drawing.Size(235, 50);
            this.Name = "CurveRangeLocator";
            this.Size = new System.Drawing.Size(235, 50);
            this.ResumeLayout(false);

        }

        private void RangeGetorXOnRangeChanged(object sender, Range range)
        {
            throw new NotImplementedException();
        }

        #endregion

        private System.Windows.Forms.Button button_srcXY;
        private RangeGetor rangeGetorY;
        private RangeGetor rangeGetorX;
    }
}
