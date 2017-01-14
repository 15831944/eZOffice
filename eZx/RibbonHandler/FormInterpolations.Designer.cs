using eZx_API.UserControls;

namespace eZx.RibbonHandler
{
    partial class FormInterpolations
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton_whatever = new System.Windows.Forms.RadioButton();
            this.radioButton_Spline = new System.Windows.Forms.RadioButton();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.buttonOk = new System.Windows.Forms.Button();
            this.rangeGetorD = new eZx_API.UserControls.RangeGetor();
            this.rangeGetorI = new eZx_API.UserControls.RangeGetor();
            this.rangeSource = new eZx_API.UserControls.CurveRangeLocator();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.radioButton_whatever);
            this.groupBox1.Controls.Add(this.radioButton_Spline);
            this.groupBox1.Location = new System.Drawing.Point(11, 112);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(231, 79);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "插值模式";
            // 
            // radioButton_whatever
            // 
            this.radioButton_whatever.AutoSize = true;
            this.radioButton_whatever.Location = new System.Drawing.Point(8, 50);
            this.radioButton_whatever.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton_whatever.Name = "radioButton_whatever";
            this.radioButton_whatever.Size = new System.Drawing.Size(41, 16);
            this.radioButton_whatever.TabIndex = 0;
            this.radioButton_whatever.Text = "***";
            this.radioButton_whatever.UseVisualStyleBackColor = true;
            // 
            // radioButton_Spline
            // 
            this.radioButton_Spline.AutoSize = true;
            this.radioButton_Spline.Checked = true;
            this.radioButton_Spline.Location = new System.Drawing.Point(8, 23);
            this.radioButton_Spline.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton_Spline.Name = "radioButton_Spline";
            this.radioButton_Spline.Size = new System.Drawing.Size(71, 16);
            this.radioButton_Spline.TabIndex = 0;
            this.radioButton_Spline.TabStop = true;
            this.radioButton_Spline.Text = "样条插值";
            this.radioButton_Spline.UseVisualStyleBackColor = true;
            // 
            // button_Cancel
            // 
            this.button_Cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Cancel.Location = new System.Drawing.Point(120, 195);
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(2);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(56, 20);
            this.button_Cancel.TabIndex = 7;
            this.button_Cancel.Text = "取消";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // buttonOk
            // 
            this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOk.Location = new System.Drawing.Point(181, 195);
            this.buttonOk.Margin = new System.Windows.Forms.Padding(2);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(56, 20);
            this.buttonOk.TabIndex = 8;
            this.buttonOk.Text = "确定";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // rangeGetorD
            // 
            this.rangeGetorD.ButtonText = "D";
            this.rangeGetorD.LabelText = "结果";
            this.rangeGetorD.Location = new System.Drawing.Point(10, 85);
            this.rangeGetorD.Name = "rangeGetorD";
            this.rangeGetorD.Size = new System.Drawing.Size(212, 22);
            this.rangeGetorD.TabIndex = 11;
            // 
            // rangeGetorI
            // 
            this.rangeGetorI.ButtonText = "I";
            this.rangeGetorI.LabelText = "插值x源";
            this.rangeGetorI.Location = new System.Drawing.Point(10, 58);
            this.rangeGetorI.Name = "rangeGetorI";
            this.rangeGetorI.Size = new System.Drawing.Size(212, 22);
            this.rangeGetorI.TabIndex = 10;
            // 
            // rangeSource
            // 
            this.rangeSource.Location = new System.Drawing.Point(6, 2);
            this.rangeSource.MaximumSize = new System.Drawing.Size(235, 50);
            this.rangeSource.MinimumSize = new System.Drawing.Size(235, 50);
            this.rangeSource.Name = "rangeSource";
            this.rangeSource.Size = new System.Drawing.Size(235, 50);
            this.rangeSource.TabIndex = 0;
            // 
            // FormInterpolations
            // 
            this.AcceptButton = this.buttonOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(248, 226);
            this.Controls.Add(this.rangeGetorD);
            this.Controls.Add(this.rangeGetorI);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.rangeSource);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "FormInterpolations";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "曲线插值";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormInterpolations_FormClosing);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormInterpolations_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private CurveRangeLocator rangeSource;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton_whatever;
        private System.Windows.Forms.RadioButton radioButton_Spline;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.Button buttonOk;
        private RangeGetor rangeGetorI;
        private RangeGetor rangeGetorD;
    }
}