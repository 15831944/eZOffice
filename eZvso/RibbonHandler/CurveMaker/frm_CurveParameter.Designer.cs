using System;
using System.Windows.Forms;
using eZstd.UserControls;

namespace eZvso.RibbonHandler.CurveMaker
{
    partial class frm_CurveParameter
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_CurveParameter));
            this.Column_X = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Y = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButton_polyline = new System.Windows.Forms.RadioButton();
            this.radioButton_nurbs = new System.Windows.Forms.RadioButton();
            this.radioButton_spline = new System.Windows.Forms.RadioButton();
            this.buttonOk = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.buttonCancel = new System.Windows.Forms.Button();
            this.dataGridView1 = new eZDataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.textBoxTolerance = new TextBoxNum();
            this.panelTolerance = new System.Windows.Forms.Panel();
            this.labelTolerance = new System.Windows.Forms.Label();
            this.label_degree = new System.Windows.Forms.Label();
            this.textBox_degree = new TextBoxNum();
            this.radioButton_bezier = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panelTolerance.SuspendLayout();
            this.SuspendLayout();
            // 
            // Column_X
            // 
            this.Column_X.HeaderText = "X";
            this.Column_X.Name = "Column_X";
            // 
            // Column_Y
            // 
            this.Column_Y.HeaderText = "Y";
            this.Column_Y.Name = "Column_Y";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(143, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(275, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "输入控制点的坐标（单位为英寸 1inch = 25.4mm）";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton_polyline);
            this.groupBox1.Controls.Add(this.radioButton_bezier);
            this.groupBox1.Controls.Add(this.radioButton_nurbs);
            this.groupBox1.Controls.Add(this.radioButton_spline);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(121, 108);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "绘制模式";
            // 
            // radioButton_polyline
            // 
            this.radioButton_polyline.AutoSize = true;
            this.radioButton_polyline.Location = new System.Drawing.Point(19, 20);
            this.radioButton_polyline.Name = "radioButton_polyline";
            this.radioButton_polyline.Size = new System.Drawing.Size(59, 16);
            this.radioButton_polyline.TabIndex = 0;
            this.radioButton_polyline.Text = "多段线";
            this.radioButton_polyline.UseVisualStyleBackColor = true;
            // 
            // radioButton_nurbs
            // 
            this.radioButton_nurbs.AutoSize = true;
            this.radioButton_nurbs.Location = new System.Drawing.Point(19, 85);
            this.radioButton_nurbs.Name = "radioButton_nurbs";
            this.radioButton_nurbs.Size = new System.Drawing.Size(77, 16);
            this.radioButton_nurbs.TabIndex = 0;
            this.radioButton_nurbs.Text = "NURBS曲线";
            this.toolTip1.SetToolTip(this.radioButton_nurbs, "非均匀有理B样条曲线（Non-Uniform Rational B-Splines）");
            this.radioButton_nurbs.UseVisualStyleBackColor = true;
            // 
            // radioButton_spline
            // 
            this.radioButton_spline.AutoSize = true;
            this.radioButton_spline.Checked = true;
            this.radioButton_spline.Location = new System.Drawing.Point(19, 41);
            this.radioButton_spline.Name = "radioButton_spline";
            this.radioButton_spline.Size = new System.Drawing.Size(71, 16);
            this.radioButton_spline.TabIndex = 0;
            this.radioButton_spline.TabStop = true;
            this.radioButton_spline.Text = "样条曲线";
            this.toolTip1.SetToolTip(this.radioButton_spline, "距离容差越大，生成的曲线与控制点的偏差越大，曲线越光滑");
            this.radioButton_spline.UseVisualStyleBackColor = true;
            // 
            // buttonOk
            // 
            this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonOk.Location = new System.Drawing.Point(337, 276);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 3;
            this.buttonOk.Text = "绘制";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Location = new System.Drawing.Point(256, 276);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 3;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Right | AnchorStyles.Left)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column_X,
            this.Column_Y});
            this.dataGridView1.Location = new System.Drawing.Point(145, 32);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(267, 238);
            this.dataGridView1.TabIndex = 4;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            // 
            // textBoxTolerance
            // 
            this.textBoxTolerance.Location = new System.Drawing.Point(4, 24);
            this.textBoxTolerance.Name = "textBoxTolerance";
            this.textBoxTolerance.Size = new System.Drawing.Size(100, 21);
            this.textBoxTolerance.TabIndex = 5;
            // 
            // panelTolerance
            // 
            this.panelTolerance.Controls.Add(this.textBox_degree);
            this.panelTolerance.Controls.Add(this.label_degree);
            this.panelTolerance.Controls.Add(this.textBoxTolerance);
            this.panelTolerance.Controls.Add(this.labelTolerance);
            this.panelTolerance.Location = new System.Drawing.Point(12, 126);
            this.panelTolerance.Name = "panelTolerance";
            this.panelTolerance.Size = new System.Drawing.Size(121, 105);
            this.panelTolerance.TabIndex = 6;
            // 
            // labelTolerance
            // 
            this.labelTolerance.AutoSize = true;
            this.labelTolerance.Location = new System.Drawing.Point(4, 6);
            this.labelTolerance.Name = "labelTolerance";
            this.labelTolerance.Size = new System.Drawing.Size(89, 12);
            this.labelTolerance.TabIndex = 1;
            this.labelTolerance.Text = "距离容差(inch)";
            // 
            // label_degree
            // 
            this.label_degree.AutoSize = true;
            this.label_degree.Location = new System.Drawing.Point(3, 57);
            this.label_degree.Name = "label_degree";
            this.label_degree.Size = new System.Drawing.Size(41, 12);
            this.label_degree.TabIndex = 1;
            this.label_degree.Text = "degree";
            // 
            // textBox_degree
            // 
            this.textBox_degree.Location = new System.Drawing.Point(3, 75);
            this.textBox_degree.Name = "textBox_degree";
            this.textBox_degree.Size = new System.Drawing.Size(100, 21);
            this.textBox_degree.TabIndex = 5;
            this.toolTip1.SetToolTip(this.textBox_degree, "贝塞尔曲线或者 Nurbs 曲线的阶次数。");
            // 
            // radioButton_bezier
            // 
            this.radioButton_bezier.AutoSize = true;
            this.radioButton_bezier.Location = new System.Drawing.Point(19, 63);
            this.radioButton_bezier.Name = "radioButton_bezier";
            this.radioButton_bezier.Size = new System.Drawing.Size(83, 16);
            this.radioButton_bezier.TabIndex = 0;
            this.radioButton_bezier.Text = "贝塞尔曲线";
            this.radioButton_bezier.UseVisualStyleBackColor = true;
            // 
            // frm_CurveParameter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 311);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panelTolerance);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(440, 350);
            this.Name = "frm_CurveParameter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "曲线绘制";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panelTolerance.ResumeLayout(false);
            this.panelTolerance.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private eZDataGridView dataGridView1;
        private TextBoxNum textBoxTolerance;
        private TextBoxNum textBox_degree;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_X;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Y;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButton_polyline;
        private System.Windows.Forms.RadioButton radioButton_nurbs;
        private System.Windows.Forms.RadioButton radioButton_spline;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Label labelTolerance;
        private System.Windows.Forms.Panel panelTolerance;
        private System.Windows.Forms.ToolTip toolTip1;
        private RadioButton radioButton_bezier;
        private Label label_degree;
    }
}