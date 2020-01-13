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
            this.PlmTab = this.Factory.CreateRibbonTab();
            this.Essentials = this.Factory.CreateRibbonGroup();
            this.GetTemplate = this.Factory.CreateRibbonButton();
            this.Validate = this.Factory.CreateRibbonButton();
            this.UploadSheet = this.Factory.CreateRibbonButton();
            this.UpdateSheet = this.Factory.CreateRibbonButton();
            this.PlmTab.SuspendLayout();
            this.Essentials.SuspendLayout();
            this.SuspendLayout();
            // 
            // PlmTab
            // 
            this.PlmTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.PlmTab.Groups.Add(this.Essentials);
            this.PlmTab.Label = "MFB-PLM";
            this.PlmTab.Name = "PlmTab";
            // 
            // Essentials
            // 
            this.Essentials.Items.Add(this.GetTemplate);
            this.Essentials.Items.Add(this.Validate);
            this.Essentials.Items.Add(this.UploadSheet);
            this.Essentials.Items.Add(this.UpdateSheet);
            this.Essentials.Label = "Essentials";
            this.Essentials.Name = "Essentials";
            // 
            // GetTemplate
            // 
            this.GetTemplate.Label = "Get Template";
            this.GetTemplate.Name = "GetTemplate";
            this.GetTemplate.SuperTip = "Generates LineSheet Template";
            this.GetTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GetTemplate_Click);
            // 
            // Validate
            // 
            this.Validate.Enabled = false;
            this.Validate.Label = "Validate";
            this.Validate.Name = "Validate";
            this.Validate.SuperTip = "Validates Data of Current Sheet";
            this.Validate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Validate_Click);
            // 
            // UploadSheet
            // 
            this.UploadSheet.Enabled = false;
            this.UploadSheet.Label = "Upload Sheet";
            this.UploadSheet.Name = "UploadSheet";
            this.UploadSheet.SuperTip = "Uploads Data of Current Sheet to the Service";
            this.UploadSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UploadSheet_Click);
            // 
            // UpdateSheet
            // 
            this.UpdateSheet.Label = "Update";
            this.UpdateSheet.Name = "UpdateSheet";
            this.UpdateSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateSheet_Click);
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
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab PlmTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Essentials;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton GetTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Validate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UploadSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton UpdateSheet;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
