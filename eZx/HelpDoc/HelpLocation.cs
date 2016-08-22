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
	/// <summary>
	/// 帮助文档的位置
	/// </summary>
	public partial class HelpLocation
	{
		public HelpLocation()
		{
			InitializeComponent();
		}
		
		private HelpLocationSettings settings1 = new HelpLocationSettings();
		
		public void HelpLocation_Load(object sender, EventArgs e)
		{
			HelpLocation with_1 = this;
			with_1.TextBox_OfficeHelp.Text = settings1.OfficeHelp;
			with_1.TextBox_ExcelHelp.Text = settings1.ExcelHelp;
		}
		
		public void Form1_FormClosing_1(object sender, 
			FormClosingEventArgs e)
		{
			// Save settings manually.
			settings1.Save();
		}
		
		public void btnOk_Click(object sender, EventArgs e)
		{
			HelpLocation with_1 = this;
			settings1.OfficeHelp = with_1.TextBox_OfficeHelp.Text;
			settings1.ExcelHelp = with_1.TextBox_ExcelHelp.Text;
			this.Close();
		}
	}
}
