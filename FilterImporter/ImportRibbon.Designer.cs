namespace OutlookAddIn1
{
    partial class ImportRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ImportRibbon()
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
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.selectButton = this.Factory.CreateRibbonButton();
            this.importButton = this.Factory.CreateRibbonButton();
            this.debug = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Gmail Filters";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            ribbonDialogLauncherImpl1.Enabled = false;
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.selectButton);
            this.group1.Items.Add(this.importButton);
            this.group1.Items.Add(this.debug);
            this.group1.Label = "Import Filters";
            this.group1.Name = "group1";
            // 
            // selectButton
            // 
            this.selectButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.selectButton.ImageName = "AutoFilterClassic";
            this.selectButton.Label = "Select Filters.xml";
            this.selectButton.Name = "selectButton";
            this.selectButton.OfficeImageId = "AutoFilterClassic";
            this.selectButton.ShowImage = true;
            this.selectButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectXml_Click);
            // 
            // importButton
            // 
            this.importButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.importButton.Label = "Start";
            this.importButton.Name = "importButton";
            this.importButton.OfficeImageId = "MacroPlay";
            this.importButton.ShowImage = true;
            this.importButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Import_Click);
            // 
            // debug
            // 
            this.debug.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.debug.Label = "Remove All Rules";
            this.debug.Name = "debug";
            this.debug.OfficeImageId = "QueryDelete";
            this.debug.ShowImage = true;
            this.debug.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Debug_Click);
            // 
            // ImportRibbon
            // 
            this.Name = "ImportRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ImportRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton debug;
    }

    partial class ThisRibbonCollection
    {
        internal ImportRibbon TestRibbon
        {
            get { return this.GetRibbon<ImportRibbon>(); }
        }
    }
}
