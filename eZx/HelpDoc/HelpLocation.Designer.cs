// VBConversions Note: VB project level imports
using System.Collections.Generic;
using System;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Linq;
// End of VB project level imports


namespace eZx
{
	partial class HelpLocation : System.Windows.Forms.Form
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
			base.Load += new System.EventHandler(HelpLocation_Load);
			base.FormClosing += new System.Windows.Forms.FormClosingEventHandler(Form1_FormClosing_1);
			this.TextBox_OfficeHelp = new System.Windows.Forms.TextBox();
			this.Label2 = new System.Windows.Forms.Label();
			this.TextBox_ExcelHelp = new System.Windows.Forms.TextBox();
			this.btnOk = new System.Windows.Forms.Button();
			this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
			this.SuspendLayout();
			//
			//Label1
			//
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(13, 13);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(179, 12);
			this.Label1.TabIndex = 0;
			this.Label1.Text = "Office VBA 帮助文档所在文件夹";
			//
			//TextBox_OfficeHelp
			//
			this.TextBox_OfficeHelp.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.TextBox_OfficeHelp.Location = new System.Drawing.Point(24, 29);
			this.TextBox_OfficeHelp.Name = "TextBox_OfficeHelp";
			this.TextBox_OfficeHelp.Size = new System.Drawing.Size(524, 21);
			this.TextBox_OfficeHelp.TabIndex = 1;
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(13, 57);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(161, 12);
			this.Label2.TabIndex = 0;
			this.Label2.Text = "Excel VBA 帮助文档所在文件";
			//
			//TextBox_ExcelHelp
			//
			this.TextBox_ExcelHelp.Anchor = (System.Windows.Forms.AnchorStyles) ((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.TextBox_ExcelHelp.Location = new System.Drawing.Point(24, 73);
			this.TextBox_ExcelHelp.Name = "TextBox_ExcelHelp";
			this.TextBox_ExcelHelp.Size = new System.Drawing.Size(524, 21);
			this.TextBox_ExcelHelp.TabIndex = 1;
			//
			//btnOk
			//
			this.btnOk.Location = new System.Drawing.Point(473, 105);
			this.btnOk.Name = "btnOk";
			this.btnOk.Size = new System.Drawing.Size(75, 23);
			this.btnOk.TabIndex = 2;
			this.btnOk.Text = "确定";
			this.btnOk.UseVisualStyleBackColor = true;
			//
			//HelpLocation
			//
			this.AcceptButton = this.btnOk;
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(560, 140);
			this.Controls.Add(this.btnOk);
			this.Controls.Add(this.TextBox_ExcelHelp);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.TextBox_OfficeHelp);
			this.Controls.Add(this.Label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "HelpLocation";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "HelpLocation";
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		
		internal Label Label1;
		internal TextBox TextBox_OfficeHelp;
		internal Label Label2;
		internal TextBox TextBox_ExcelHelp;
		internal Button btnOk;
	}
	
}
