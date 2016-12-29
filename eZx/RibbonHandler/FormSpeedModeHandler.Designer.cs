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
            this.textBox_srcX = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_srcY = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.button_srcX = new System.Windows.Forms.Button();
            this.button_srcY = new System.Windows.Forms.Button();
            this.button_srcXY = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.radioButton_PointCount = new System.Windows.Forms.RadioButton();
            this.radioButton_XSegment = new System.Windows.Forms.RadioButton();
            this.textBox_srcD = new System.Windows.Forms.TextBox();
            this.button_srcD = new System.Windows.Forms.Button();
            this.button_Ok = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.numericUpDown_PointsSegments = new System.Windows.Forms.NumericUpDown();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_PointsSegments)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox_srcX
            // 
            this.textBox_srcX.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_srcX.Location = new System.Drawing.Point(64, 10);
            this.textBox_srcX.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBox_srcX.Name = "textBox_srcX";
            this.textBox_srcX.Size = new System.Drawing.Size(134, 21);
            this.textBox_srcX.TabIndex = 0;
            this.textBox_srcX.Enter += new System.EventHandler(this.textBox_Datasource_Enter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 14);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "X数据源";
            // 
            // textBox_srcY
            // 
            this.textBox_srcY.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_srcY.Location = new System.Drawing.Point(64, 34);
            this.textBox_srcY.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBox_srcY.Name = "textBox_srcY";
            this.textBox_srcY.Size = new System.Drawing.Size(134, 21);
            this.textBox_srcY.TabIndex = 1;
            this.textBox_srcY.Enter += new System.EventHandler(this.textBox_Datasource_Enter);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 38);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "Y数据源";
            // 
            // button_srcX
            // 
            this.button_srcX.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_srcX.Location = new System.Drawing.Point(201, 10);
            this.button_srcX.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button_srcX.Name = "button_srcX";
            this.button_srcX.Size = new System.Drawing.Size(17, 20);
            this.button_srcX.TabIndex = 3;
            this.button_srcX.Text = "X";
            this.button_srcX.UseVisualStyleBackColor = true;
            this.button_srcX.Click += new System.EventHandler(this.button_Datasource_Click);
            // 
            // button_srcY
            // 
            this.button_srcY.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_srcY.Location = new System.Drawing.Point(200, 34);
            this.button_srcY.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button_srcY.Name = "button_srcY";
            this.button_srcY.Size = new System.Drawing.Size(17, 20);
            this.button_srcY.TabIndex = 4;
            this.button_srcY.Text = "Y";
            this.button_srcY.UseVisualStyleBackColor = true;
            this.button_srcY.Click += new System.EventHandler(this.button_Datasource_Click);
            // 
            // button_srcXY
            // 
            this.button_srcXY.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_srcXY.Location = new System.Drawing.Point(223, 10);
            this.button_srcXY.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button_srcXY.Name = "button_srcXY";
            this.button_srcXY.Size = new System.Drawing.Size(17, 45);
            this.button_srcXY.TabIndex = 5;
            this.button_srcXY.Text = "XY";
            this.button_srcXY.UseVisualStyleBackColor = true;
            this.button_srcXY.Click += new System.EventHandler(this.button_Datasource_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 64);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "目标位置";
            this.toolTip1.SetToolTip(this.label3, "缩减后的XY数据列的左上角位置");
            // 
            // radioButton_PointCount
            // 
            this.radioButton_PointCount.AutoSize = true;
            this.radioButton_PointCount.Checked = true;
            this.radioButton_PointCount.Location = new System.Drawing.Point(8, 23);
            this.radioButton_PointCount.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            this.radioButton_XSegment.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.radioButton_XSegment.Name = "radioButton_XSegment";
            this.radioButton_XSegment.Size = new System.Drawing.Size(65, 16);
            this.radioButton_XSegment.TabIndex = 0;
            this.radioButton_XSegment.Text = "X轴分段";
            this.toolTip1.SetToolTip(this.radioButton_XSegment, "考虑X轴的分布疏密，将X轴所占区分均分为数断");
            this.radioButton_XSegment.UseVisualStyleBackColor = true;
            // 
            // textBox_srcD
            // 
            this.textBox_srcD.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_srcD.Location = new System.Drawing.Point(64, 59);
            this.textBox_srcD.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.textBox_srcD.Name = "textBox_srcD";
            this.textBox_srcD.Size = new System.Drawing.Size(134, 21);
            this.textBox_srcD.TabIndex = 2;
            this.textBox_srcD.Enter += new System.EventHandler(this.textBox_Datasource_Enter);
            // 
            // button_srcD
            // 
            this.button_srcD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button_srcD.Location = new System.Drawing.Point(200, 59);
            this.button_srcD.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button_srcD.Name = "button_srcD";
            this.button_srcD.Size = new System.Drawing.Size(17, 20);
            this.button_srcD.TabIndex = 6;
            this.button_srcD.Text = "D";
            this.button_srcD.UseVisualStyleBackColor = true;
            this.button_srcD.Click += new System.EventHandler(this.button_Datasource_Click);
            // 
            // button_Ok
            // 
            this.button_Ok.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Ok.Location = new System.Drawing.Point(184, 190);
            this.button_Ok.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            // FormSpeedModeHandler
            // 
            this.AcceptButton = this.button_Ok;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(250, 225);
            this.Controls.Add(this.numericUpDown_PointsSegments);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.button_Ok);
            this.Controls.Add(this.button_srcXY);
            this.Controls.Add(this.button_srcD);
            this.Controls.Add(this.button_srcY);
            this.Controls.Add(this.button_srcX);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox_srcD);
            this.Controls.Add(this.textBox_srcY);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox_srcX);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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

        private System.Windows.Forms.TextBox textBox_srcX;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_srcY;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button_srcX;
        private System.Windows.Forms.Button button_srcY;
        private System.Windows.Forms.Button button_srcXY;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox textBox_srcD;
        private System.Windows.Forms.Button button_srcD;
        private System.Windows.Forms.Button button_Ok;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton_XSegment;
        private System.Windows.Forms.RadioButton radioButton_PointCount;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown numericUpDown_PointsSegments;
    }
}