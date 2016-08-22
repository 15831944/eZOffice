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


namespace eZvso
{
	partial class Dialog_CircleArray : System.Windows.Forms.Form
	{
		
		//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && components != null)
				{
					components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}
		
		//Required by the Windows Form Designer
		private System.ComponentModel.Container components = null;
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]private void InitializeComponent()
		{
			this.Label1 = new System.Windows.Forms.Label();
			this.txtAngle = new System.Windows.Forms.TextBox();
			this.txtAngle.TextChanged += new System.EventHandler(this.txtAngle_TextChanged);
			this.Label2 = new System.Windows.Forms.Label();
			this.txtNum = new System.Windows.Forms.TextBox();
			this.txtNum.TextChanged += new System.EventHandler(this.txtNum_TextChanged);
			this.CheckBox_preserveDirection = new System.Windows.Forms.CheckBox();
			this.CheckBox_preserveDirection.CheckedChanged += new System.EventHandler(this.CheckBox_preserveDirection_CheckedChanged);
			this.Label3 = new System.Windows.Forms.Label();
		
			this.Label4 = new System.Windows.Forms.Label();
			this.RadioButton_Center = new System.Windows.Forms.RadioButton();
			this.RadioButton_Center.CheckedChanged += new System.EventHandler(this.RadioButton_Center_CheckedChanged);
			this.RadioButton_Border = new System.Windows.Forms.RadioButton();
			this.RadioButton_Border.CheckedChanged += new System.EventHandler(this.RadioButton_Center_CheckedChanged);
			this.btnOK = new System.Windows.Forms.Button();
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			this.btnCancel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			//
			//Label1
			//
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(24, 34);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(47, 12);
			this.Label1.TabIndex = 0;
			this.Label1.Text = "角度 ：";
			//
			//txtAngle
			//
			this.txtAngle.Location = new System.Drawing.Point(78, 31);
			this.txtAngle.Name = "txtAngle";
			this.txtAngle.Size = new System.Drawing.Size(100, 21);
			this.txtAngle.TabIndex = 1;
			this.txtAngle.Text = "360";
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(12, 61);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(59, 12);
			this.Label2.TabIndex = 0;
			this.Label2.Text = "总数量 ：";
			//
			//txtNum
			//
			this.txtNum.Location = new System.Drawing.Point(78, 58);
			this.txtNum.Name = "txtNum";
			this.txtNum.Size = new System.Drawing.Size(100, 21);
			this.txtNum.TabIndex = 1;
			this.txtNum.Text = "4";
			//
			//CheckBox_preserveDirection
			//
			this.CheckBox_preserveDirection.AutoSize = true;
			this.CheckBox_preserveDirection.Location = new System.Drawing.Point(26, 165);
			this.CheckBox_preserveDirection.Name = "CheckBox_preserveDirection";
			this.CheckBox_preserveDirection.Size = new System.Drawing.Size(174, 16);
			this.CheckBox_preserveDirection.TabIndex = 2;
			this.CheckBox_preserveDirection.Text = "与主形状的旋转保持一致(&M)";
			this.CheckBox_preserveDirection.UseVisualStyleBackColor = true;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(13, 13);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(29, 12);
			this.Label3.TabIndex = 3;
			this.Label3.Text = "布局";
			//
			
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(12, 93);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(29, 12);
			this.Label4.TabIndex = 3;
			this.Label4.Text = "间距";
			//
			//RadioButton_Center
			//
			this.RadioButton_Center.AutoSize = true;
			this.RadioButton_Center.Checked = true;
			this.RadioButton_Center.Location = new System.Drawing.Point(26, 119);
			this.RadioButton_Center.Name = "RadioButton_Center";
			this.RadioButton_Center.Size = new System.Drawing.Size(113, 16);
			this.RadioButton_Center.TabIndex = 5;
			this.RadioButton_Center.TabStop = true;
			this.RadioButton_Center.Text = "形状中心之间(&S)";
			this.RadioButton_Center.UseVisualStyleBackColor = true;
			//
			//RadioButton_Border
			//
			this.RadioButton_Border.AutoSize = true;
			this.RadioButton_Border.Enabled = false;
			this.RadioButton_Border.Location = new System.Drawing.Point(26, 141);
			this.RadioButton_Border.Name = "RadioButton_Border";
			this.RadioButton_Border.Size = new System.Drawing.Size(113, 16);
			this.RadioButton_Border.TabIndex = 5;
			this.RadioButton_Border.TabStop = true;
			this.RadioButton_Border.Text = "形状边缘之间(&E)";
			this.RadioButton_Border.UseVisualStyleBackColor = true;
			//
			//btnOK
			//
			this.btnOK.Location = new System.Drawing.Point(116, 199);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(75, 23);
			this.btnOK.TabIndex = 6;
			this.btnOK.Text = "确定";
			this.btnOK.UseVisualStyleBackColor = true;
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(197, 199);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(75, 23);
			this.btnCancel.TabIndex = 6;
			this.btnCancel.Text = "取消";
			this.btnCancel.UseVisualStyleBackColor = true;
			//
			//Dialog_CircleArray
			//
			this.AcceptButton = this.btnOK;
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(284, 231);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.RadioButton_Border);
			this.Controls.Add(this.RadioButton_Center);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.CheckBox_preserveDirection);
			this.Controls.Add(this.txtNum);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.txtAngle);
			this.Controls.Add(this.Label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "Dialog_CircleArray";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "旋转阵列";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.TextBox txtAngle;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.TextBox txtNum;
		internal System.Windows.Forms.CheckBox CheckBox_preserveDirection;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.RadioButton RadioButton_Center;
		internal System.Windows.Forms.RadioButton RadioButton_Border;
		internal System.Windows.Forms.Button btnOK;
		internal System.Windows.Forms.Button btnCancel;
	}
	
}
