namespace MyntraExcelAddin
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            this.PlmTab = this.Factory.CreateRibbonTab();
            this.Essentials = this.Factory.CreateRibbonGroup();
            this.GetTemplate = this.Factory.CreateRibbonButton();
            this.Validate = this.Factory.CreateRibbonButton();
            this.UploadSheet = this.Factory.CreateRibbonButton();
            this.UpdateSheet = this.Factory.CreateRibbonButton();
            this.Browser = this.Factory.CreateRibbonGroup();
            this.BrowserBack = this.Factory.CreateRibbonButton();
            this.BrowserStart = this.Factory.CreateRibbonButton();
            this.BrowserStop = this.Factory.CreateRibbonButton();
            this.BrowserNext = this.Factory.CreateRibbonButton();
            this.PlmTab.SuspendLayout();
            this.Essentials.SuspendLayout();
            this.Browser.SuspendLayout();
            this.SuspendLayout();
            // 
            // PlmTab
            // 
            this.PlmTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PlmTab.Groups.Add(this.Essentials);
            this.PlmTab.Groups.Add(this.Browser);
            this.PlmTab.Label = "MFB-PLM";
            this.PlmTab.Name = "PlmTab";
            // 
            // Essentials
            // 
            this.Essentials.Items.Add(this.GetTemplate);
            this.Essentials.Items.Add(this.Validate);
            this.Essentials.Items.Add(this.UploadSheet);
            this.Essentials.Items.Add(this.UpdateSheet);
            this.Essentials.Label = "LineSheet";
            this.Essentials.Name = "Essentials";
            // 
            // GetTemplate
            // 
            this.GetTemplate.Image = ((System.Drawing.Image)(resources.GetObject("GetTemplate.Image")));
            this.GetTemplate.Label = "Get Template";
            this.GetTemplate.Name = "GetTemplate";
            this.GetTemplate.ShowImage = true;
            this.GetTemplate.SuperTip = "Generates LineSheet Template";
            this.GetTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetTemplate_Click);
            // 
            // Validate
            // 
            this.Validate.Enabled = false;
            this.Validate.Image = ((System.Drawing.Image)(resources.GetObject("Validate.Image")));
            this.Validate.Label = "Validate";
            this.Validate.Name = "Validate";
            this.Validate.ShowImage = true;
            this.Validate.SuperTip = "Validates Data of Current Sheet";
            this.Validate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Validate_Click);
            // 
            // UploadSheet
            // 
            this.UploadSheet.Enabled = false;
            this.UploadSheet.Image = ((System.Drawing.Image)(resources.GetObject("UploadSheet.Image")));
            this.UploadSheet.Label = "Upload Sheet";
            this.UploadSheet.Name = "UploadSheet";
            this.UploadSheet.ShowImage = true;
            this.UploadSheet.SuperTip = "Uploads Data of Current Sheet to the Service";
            this.UploadSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadSheet_Click);
            // 
            // UpdateSheet
            // 
            this.UpdateSheet.Image = ((System.Drawing.Image)(resources.GetObject("UpdateSheet.Image")));
            this.UpdateSheet.Label = "Update";
            this.UpdateSheet.Name = "UpdateSheet";
            this.UpdateSheet.ShowImage = true;
            this.UpdateSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateSheet_Click);
            // 
            // Browser
            // 
            this.Browser.Items.Add(this.BrowserBack);
            this.Browser.Items.Add(this.BrowserStart);
            this.Browser.Items.Add(this.BrowserStop);
            this.Browser.Items.Add(this.BrowserNext);
            this.Browser.Label = "Handover Browser";
            this.Browser.Name = "Browser";
            // 
            // BrowserBack
            // 
            this.BrowserBack.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowserBack.Image = ((System.Drawing.Image)(resources.GetObject("BrowserBack.Image")));
            this.BrowserBack.Label = "Back";
            this.BrowserBack.Name = "BrowserBack";
            this.BrowserBack.ShowImage = true;
            // 
            // BrowserStart
            // 
            this.BrowserStart.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowserStart.Image = ((System.Drawing.Image)(resources.GetObject("BrowserStart.Image")));
            this.BrowserStart.Label = "Start";
            this.BrowserStart.Name = "BrowserStart";
            this.BrowserStart.ShowImage = true;
            // 
            // BrowserStop
            // 
            this.BrowserStop.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowserStop.Image = ((System.Drawing.Image)(resources.GetObject("BrowserStop.Image")));
            this.BrowserStop.Label = "Stop";
            this.BrowserStop.Name = "BrowserStop";
            this.BrowserStop.ShowImage = true;
            // 
            // BrowserNext
            // 
            this.BrowserNext.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BrowserNext.Image = ((System.Drawing.Image)(resources.GetObject("BrowserNext.Image")));
            this.BrowserNext.Label = "Next";
            this.BrowserNext.Name = "BrowserNext";
            this.BrowserNext.ShowImage = true;
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.PlmTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.PlmTab.ResumeLayout(false);
            this.PlmTab.PerformLayout();
            this.Essentials.ResumeLayout(false);
            this.Essentials.PerformLayout();
            this.Browser.ResumeLayout(false);
            this.Browser.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PlmTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Essentials;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Validate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UploadSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Browser;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowserBack;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowserStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowserStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BrowserNext;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
