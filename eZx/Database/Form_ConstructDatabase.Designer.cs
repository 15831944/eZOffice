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


namespace eZx.Database
{
	partial class Form_ConstructDatabase : System.Windows.Forms.Form
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
			this.components = new System.ComponentModel.Container();
			this.Load += new System.EventHandler(Form_ConstructDatabase_Load);
			List_FieldInfo.AddingNew += new System.ComponentModel.AddingNewEventHandler(FieldInfo_AddingNew);
			this.Label1 = new System.Windows.Forms.Label();
			this.Label2 = new System.Windows.Forms.Label();
			this.Label3 = new System.Windows.Forms.Label();
			this.ComboBox_FieldType = new System.Windows.Forms.ComboBox();
			this.ComboBox_FieldType.SelectedValueChanged += new System.EventHandler(this.ChangeFieldType);
			this.ComboBox_CommonDataType = new System.Windows.Forms.ComboBox();
			this.ComboBox_CommonDataType.SelectedValueChanged += new System.EventHandler(this.ChangeAllFieldDataType);
			this.eZDataGridView1 = new eZstd.UserControl.eZDataGridView();
			this.eZDataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.eZDataGridView1_CellContentClick);
			this.eZDataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.eZDataGridView1_DataError);
			this.Label4 = new System.Windows.Forms.Label();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.CheckBox1 = new System.Windows.Forms.CheckBox();
			this.btnCheckAllFields = new System.Windows.Forms.Button();
			this.btnCheckAllFields.Click += new System.EventHandler(this.CheckAllFields);
			this.btnCancel = new System.Windows.Forms.Button();
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			this.Label5 = new System.Windows.Forms.Label();
			this.CheckBox2 = new System.Windows.Forms.CheckBox();
			this.CheckBox2.CheckedChanged += new System.EventHandler(this.CheckBox2_CheckedChanged);
			((System.ComponentModel.ISupportInitialize) this.eZDataGridView1).BeginInit();
			this.SuspendLayout();
			//
			//Label1
			//
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(13, 16);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(53, 12);
			this.Label1.TabIndex = 0;
			this.Label1.Text = "表头类型";
			//
			//Label2
			//
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(13, 34);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(0, 12);
			this.Label2.TabIndex = 1;
			//
			//Label3
			//
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(13, 49);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(77, 12);
			this.Label3.TabIndex = 2;
			this.Label3.Text = "基本数据类型";
			//
			//ComboBox_FieldType
			//
			this.ComboBox_FieldType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_FieldType.FormattingEnabled = true;
			this.ComboBox_FieldType.Location = new System.Drawing.Point(117, 12);
			this.ComboBox_FieldType.Name = "ComboBox_FieldType";
			this.ComboBox_FieldType.Size = new System.Drawing.Size(88, 20);
			this.ComboBox_FieldType.TabIndex = 5;
			this.ToolTip1.SetToolTip(this.ComboBox_FieldType, "从第二个字段开始，字段的名称本身的数据类型");
			//
			//ComboBox_CommonDataType
			//
			this.ComboBox_CommonDataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ComboBox_CommonDataType.FormattingEnabled = true;
			this.ComboBox_CommonDataType.Location = new System.Drawing.Point(117, 46);
			this.ComboBox_CommonDataType.Name = "ComboBox_CommonDataType";
			this.ComboBox_CommonDataType.Size = new System.Drawing.Size(88, 20);
			this.ComboBox_CommonDataType.TabIndex = 6;
			this.ToolTip1.SetToolTip(this.ComboBox_CommonDataType, "所有字段的数据类型");
			//
			//eZDataGridView1
			//
			this.eZDataGridView1.Anchor = (System.Windows.Forms.AnchorStyles) (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right);
			this.eZDataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.eZDataGridView1.Location = new System.Drawing.Point(14, 142);
			this.eZDataGridView1.Name = "eZDataGridView1";
			this.eZDataGridView1.RowTemplate.Height = 23;
			this.eZDataGridView1.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.False;
			this.eZDataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.eZDataGridView1.Size = new System.Drawing.Size(407, 197);
			this.eZDataGridView1.TabIndex = 7;
			//
			//Label4
			//
			this.Label4.AutoSize = true;
			this.Label4.Location = new System.Drawing.Point(12, 118);
			this.Label4.Name = "Label4";
			this.Label4.Size = new System.Drawing.Size(65, 12);
			this.Label4.TabIndex = 8;
			this.Label4.Text = "各字段信息";
			//
			//
			//CheckBox1
			//
			this.CheckBox1.AutoSize = true;
			this.CheckBox1.Checked = true;
			this.CheckBox1.CheckState = System.Windows.Forms.CheckState.Indeterminate;
			this.CheckBox1.Location = new System.Drawing.Point(235, 15);
			this.CheckBox1.Name = "CheckBox1";
			this.CheckBox1.Size = new System.Drawing.Size(48, 16);
			this.CheckBox1.TabIndex = 11;
			this.CheckBox1.Text = "升序";
			this.CheckBox1.ThreeState = true;
			this.ToolTip1.SetToolTip(this.CheckBox1, "如果为Checked，则为升序；如果为UnChecked，则为降序；如果为Indeterminate，则不考虑排序。");
			this.CheckBox1.UseVisualStyleBackColor = true;
			//
			//btnCheckAllFields
			//
			this.btnCheckAllFields.Location = new System.Drawing.Point(346, 15);
			this.btnCheckAllFields.Name = "btnCheckAllFields";
			this.btnCheckAllFields.Size = new System.Drawing.Size(75, 23);
			this.btnCheckAllFields.TabIndex = 10;
			this.btnCheckAllFields.Text = "数据验证";
			this.btnCheckAllFields.UseVisualStyleBackColor = true;
			//
			//btnCancel
			//
			this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnCancel.Location = new System.Drawing.Point(346, 49);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(75, 23);
			this.btnCancel.TabIndex = 12;
			this.btnCancel.Text = "取消";
			this.btnCancel.UseVisualStyleBackColor = true;
			//
			//Label5
			//
			this.Label5.AutoSize = true;
			this.Label5.Location = new System.Drawing.Point(13, 84);
			this.Label5.Name = "Label5";
			this.Label5.Size = new System.Drawing.Size(53, 12);
			this.Label5.TabIndex = 13;
			this.Label5.Text = "允许空值";
			//
			//CheckBox2
			//
			this.CheckBox2.AutoSize = true;
			this.CheckBox2.Checked = true;
			this.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked;
			this.CheckBox2.Location = new System.Drawing.Point(117, 83);
			this.CheckBox2.Name = "CheckBox2";
			this.CheckBox2.Size = new System.Drawing.Size(72, 16);
			this.CheckBox2.TabIndex = 14;
			this.CheckBox2.Text = "允许空值";
			this.CheckBox2.UseVisualStyleBackColor = true;
			//
			//Form_ConstructDatabase
			//
			this.AcceptButton = this.btnCheckAllFields;
			this.AutoScaleDimensions = new System.Drawing.SizeF((float) (6.0F), (float) (12.0F));
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnCancel;
			this.ClientSize = new System.Drawing.Size(434, 351);
			this.Controls.Add(this.CheckBox2);
			this.Controls.Add(this.Label5);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.CheckBox1);
			this.Controls.Add(this.btnCheckAllFields);
			this.Controls.Add(this.Label4);
			this.Controls.Add(this.eZDataGridView1);
			this.Controls.Add(this.ComboBox_CommonDataType);
			this.Controls.Add(this.ComboBox_FieldType);
			this.Controls.Add(this.Label3);
			this.Controls.Add(this.Label2);
			this.Controls.Add(this.Label1);
			this.Name = "Form_ConstructDatabase";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "构造数据库";
			((System.ComponentModel.ISupportInitialize) this.eZDataGridView1).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();
			
		}
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.ComboBox ComboBox_FieldType;
		internal System.Windows.Forms.ComboBox ComboBox_CommonDataType;
		internal eZstd.UserControl.eZDataGridView eZDataGridView1;
		internal System.Windows.Forms.Label Label4;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.Button btnCheckAllFields;
		internal System.Windows.Forms.CheckBox CheckBox1;
		internal System.Windows.Forms.Button btnCancel;
		internal System.Windows.Forms.Label Label5;
		internal System.Windows.Forms.CheckBox CheckBox2;
	}
	
}
