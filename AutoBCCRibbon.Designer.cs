namespace AutoBCC
{
    partial class AutoBCCRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AutoBCCRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.EnableToggle = this.Factory.CreateRibbonToggleButton();
            this.MailText = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Auto BCC";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.EnableToggle);
            this.group1.Items.Add(this.MailText);
            this.group1.Label = "Configuration";
            this.group1.Name = "group1";
            // 
            // EnableToggle
            // 
            this.EnableToggle.Checked = global::AutoBCC.Properties.Settings.Default.AddInEnabled;
            this.EnableToggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.EnableToggle.Image = global::AutoBCC.Properties.Resources.toggle_on_solid;
            this.EnableToggle.Label = "Ein / Aus";
            this.EnableToggle.Name = "EnableToggle";
            this.EnableToggle.ShowImage = true;
            this.EnableToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EnableToggle_Click);
            // 
            // MailText
            // 
            this.MailText.Enabled = global::AutoBCC.Properties.Settings.Default.AddInEnabled;
            this.MailText.Image = global::AutoBCC.Properties.Resources.envelope_regular;
            this.MailText.Label = "BCC E-Mail";
            this.MailText.Name = "MailText";
            this.MailText.ShowImage = true;
            this.MailText.SizeString = "012345678901234567890132456789";
            this.MailText.Text = global::AutoBCC.Properties.Settings.Default.MailText;
            this.MailText.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.MailText_TextChanged);
            // 
            // AutoBCCRibbon
            // 
            this.Name = "AutoBCCRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton EnableToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox MailText;
    }

    partial class ThisRibbonCollection
    {
        internal AutoBCCRibbon MyRibbon
        {
            get { return this.GetRibbon<AutoBCCRibbon>(); }
        }
    }
}
