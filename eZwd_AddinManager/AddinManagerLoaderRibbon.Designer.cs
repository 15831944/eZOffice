﻿using eZwd_AddinManager;

namespace eZwd_AddinManager
{
    partial class AddinManagerLoaderRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AddinManagerLoaderRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddinManagerLoaderRibbon));
            this.WdAddinManager = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonAddinManager = this.Factory.CreateRibbonButton();
            this.buttonLastCommand = this.Factory.CreateRibbonButton();
            this.WdAddinManager.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // WdAddinManager
            // 
            this.WdAddinManager.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.WdAddinManager.ControlId.OfficeId = "TabDeveloper";
            this.WdAddinManager.Groups.Add(this.group1);
            this.WdAddinManager.Label = "TabDeveloper";
            this.WdAddinManager.Name = "WdAddinManager";
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
            this.buttonAddinManager.Image = global::eZwd_AddinManager.Properties.Resources.AddinManager;
            this.buttonAddinManager.Label = "Manager";
            this.buttonAddinManager.Name = "buttonAddinManager";
            this.buttonAddinManager.ScreenTip = "Word AddinManager 快速调试插件";
            this.buttonAddinManager.ShowImage = true;
            this.buttonAddinManager.SuperTip = resources.GetString("buttonAddinManager.SuperTip");
            this.buttonAddinManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddinManager_Click);
            // 
            // buttonLastCommand
            // 
            this.buttonLastCommand.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonLastCommand.Image = global::eZwd_AddinManager.Properties.Resources.LastExternalCommand;
            this.buttonLastCommand.Label = "Last";
            this.buttonLastCommand.Name = "buttonLastCommand";
            this.buttonLastCommand.ScreenTip = "上次执行的外部命令";
            this.buttonLastCommand.ShowImage = true;
            this.buttonLastCommand.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLastCommand_Click);
            // 
            // AddinManagerLoaderRibbon
            // 
            this.Name = "AddinManagerLoaderRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.WdAddinManager);
            this.Close += new System.EventHandler(this.AddinManagerLoaderRibbon_Close);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AddinManagerLoader_Load);
            this.WdAddinManager.ResumeLayout(false);
            this.WdAddinManager.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab WdAddinManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddinManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLastCommand;
    }

    partial class ThisRibbonCollection
    {
        internal AddinManagerLoaderRibbon AddinManagerLoaderRibbon
        {
            get { return this.GetRibbon<AddinManagerLoaderRibbon>(); }
        }
    }
}
