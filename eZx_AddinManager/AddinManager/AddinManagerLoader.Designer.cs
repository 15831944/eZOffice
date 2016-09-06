using eZx_AddinManager;

namespace eZx_AddinManager
{
    partial class AddinManagerLoader : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddinManagerLoader()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinManagerLoader));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonAddinManager = this.Factory.CreateRibbonButton();
            this.buttonLastCommand = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabDeveloper";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabDeveloper";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonAddinManager);
            this.group1.Items.Add(this.buttonLastCommand);
            this.group1.Label = "快速调试";
            this.group1.Name = "group1";
            // 
            // buttonAddinManager
            // 
            this.buttonAddinManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddinManager.Image = global::eZx.Properties.Resources.AddinManager;
            this.buttonAddinManager.Label = "Manager";
            this.buttonAddinManager.Name = "buttonAddinManager";
            this.buttonAddinManager.ScreenTip = "Excel AddinManager 快速调试插件";
            this.buttonAddinManager.ShowImage = true;
            this.buttonAddinManager.SuperTip = resources.GetString("buttonAddinManager.SuperTip");
            this.buttonAddinManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddinManager_Click);
            // 
            // buttonLastCommand
            // 
            this.buttonLastCommand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonLastCommand.Image = global::eZx.Properties.Resources.LastExternalCommand;
            this.buttonLastCommand.Label = "Last";
            this.buttonLastCommand.Name = "buttonLastCommand";
            this.buttonLastCommand.ScreenTip = "上次执行的外部命令";
            this.buttonLastCommand.ShowImage = true;
            this.buttonLastCommand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLastCommand_Click);
            // 
            // AddinManagerLoader
            // 
            this.Name = "AddinManagerLoader";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddinManagerLoader_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddinManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLastCommand;
    }

    partial class ThisRibbonCollection
    {
        internal AddinManagerLoader AddinManagerLoader
        {
            get { return this.GetRibbon<AddinManagerLoader>(); }
        }
    }
}
