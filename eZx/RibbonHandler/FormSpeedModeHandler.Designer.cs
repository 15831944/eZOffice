namespace eZx.RibbonHandler
{
    partial class FormSpeedModeHandler
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
            this.components = new System.ComponentModel.Container();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.radioButton_PointCount = new System.Windows.Forms.RadioButton();
            this.radioButton_XSegment = new System.Windows.Forms.RadioButton();
            this.button_Ok = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.numericUpDown_PointsSegments = new System.Windows.Forms.NumericUpDown();
            this.rangeSource = new eZx_API.UserControls.CurveRangeLocator();
            this.rangeGetorD = new eZx_API.UserControls.RangeGetor();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_PointsSegments)).BeginInit();
            this.SuspendLayout();
            // 
            // radioButton_PointCount
            // 
            this.radioButton_PointCount.AutoSize = true;
            this.radioButton_PointCount.Checked = true;
            this.radioButton_PointCount.Location = new System.Drawing.Point(8, 23);
            this.radioButton_PointCount.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton_PointCount.Name = "radioButton_PointCount";
            this.radioButton_PointCount.Size = new System.Drawing.Size(83, 16);
            this.radioButton_PointCount.TabIndex = 0;
            this.radioButton_PointCount.TabStop = true;
            this.radioButton_PointCount.Text = "数据点个数";
            this.toolTip1.SetToolTip(this.radioButton_PointCount, "不考虑XY的值，只按数据点的个数进行缩减");
            this.radioButton_PointCount.UseVisualStyleBackColor = true;
            // 
            // radioButton_XSegment
            // 
            this.radioButton_XSegment.AutoSize = true;
            this.radioButton_XSegment.Location = new System.Drawing.Point(8, 50);
            this.radioButton_XSegment.Margin = new System.Windows.Forms.Padding(2);
            this.radioButton_XSegment.Name = "radioButton_XSegment";
            this.radioButton_XSegment.Size = new System.Drawing.Size(65, 16);
            this.radioButton_XSegment.TabIndex = 0;
            this.radioButton_XSegment.Text = "X轴分段";
            this.toolTip1.SetToolTip(this.radioButton_XSegment, "考虑X轴的分布疏密，将X轴所占区分均分为数断");
            this.radioButton_XSegment.UseVisualStyleBackColor = true;
            // 
            // button_Ok
            // 
            this.button_Ok.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Ok.Location = new System.Drawing.Point(184, 190);
            this.button_Ok.Margin = new System.Windows.Forms.Padding(2);
            this.button_Ok.Name = "button_Ok";
            this.button_Ok.Size = new System.Drawing.Size(56, 20);
            this.button_Ok.TabIndex = 3;
            this.button_Ok.Text = "确定";
            this.button_Ok.UseVisualStyleBackColor = true;
            this.button_Ok.Click += new System.EventHandler(this.button_Ok_Click);
            // 
            // button_Cancel
            // 
            this.button_Cancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Cancel.Location = new System.Drawing.Point(123, 190);
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(2);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(56, 20);
            this.button_Cancel.TabIndex = 3;
            this.button_Cancel.Text = "取消";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.radioButton_XSegment);
            this.groupBox1.Controls.Add(this.radioButton_PointCount);
            this.groupBox1.Location = new System.Drawing.Point(11, 93);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(229, 91);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "缩减模式";
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(65, 194);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 12);
            this.label4.TabIndex = 1;
            this.label4.Text = "点/段";
            // 
            // numericUpDown_PointsSegments
            // 
            this.numericUpDown_PointsSegments.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numericUpDown_PointsSegments.Location = new System.Drawing.Point(11, 189);
            this.numericUpDown_PointsSegments.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numericUpDown_PointsSegments.Name = "numericUpDown_PointsSegments";
            this.numericUpDown_PointsSegments.Size = new System.Drawing.Size(51, 21);
            this.numericUpDown_PointsSegments.TabIndex = 1;
            this.numericUpDown_PointsSegments.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // rangeSource
            // 
            this.rangeSource.Location = new System.Drawing.Point(5, 4);
            this.rangeSource.MaximumSize = new System.Drawing.Size(235, 50);
            this.rangeSource.MinimumSize = new System.Drawing.Size(235, 50);
            this.rangeSource.Name = "rangeSource";
            this.rangeSource.Size = new System.Drawing.Size(235, 50);
            this.rangeSource.TabIndex = 7;
            // 
            // rangeGetorD
            // 
            this.rangeGetorD.ButtonText = "D";
            this.rangeGetorD.LabelText = "目标位置";
            this.rangeGetorD.Location = new System.Drawing.Point(9, 59);
            this.rangeGetorD.Name = "rangeGetorD";
            this.rangeGetorD.Size = new System.Drawing.Size(212, 23);
            this.rangeGetorD.TabIndex = 8;
            // 
            // FormSpeedModeHandler
            // 
            this.AcceptButton = this.button_Ok;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(250, 225);
            this.Controls.Add(this.rangeGetorD);
            this.Controls.Add(this.rangeSource);
            this.Controls.Add(this.numericUpDown_PointsSegments);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.button_Ok);
            this.Controls.Add(this.label4);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimumSize = new System.Drawing.Size(266, 264);
            this.Name = "FormSpeedModeHandler";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SpeedMode";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormSpeedModeHandler_FormClosing);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.FormSpeedModeHandler_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_PointsSegments)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button_Ok;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton_XSegment;
        private System.Windows.Forms.RadioButton radioButton_PointCount;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numericUpDown_PointsSegments;
        private eZx_API.UserControls.CurveRangeLocator rangeSource;
        private eZx_API.UserControls.RangeGetor rangeGetorD;
    }
}