namespace OpenAI
{
    partial class OpenAIRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OpenAIRibbon()
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
            this.openAITab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.respondButton = this.Factory.CreateRibbonButton();
            this.openAITab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openAITab
            // 
            this.openAITab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.openAITab.Groups.Add(this.group1);
            this.openAITab.Label = "OpenAI Tools";
            this.openAITab.Name = "openAITab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.respondButton);
            this.group1.Label = "OpenAI Tools";
            this.group1.Name = "group1";
            // 
            // respondButton
            // 
            this.respondButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.respondButton.Image = global::OpenAI.Properties.Resources.NetizineBot;
            this.respondButton.Label = "Respond";
            this.respondButton.Name = "respondButton";
            this.respondButton.ShowImage = true;
            this.respondButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OpenAIRespond_Click);
            // 
            // OpenAIRibbon
            // 
            this.Name = "OpenAIRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.openAITab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OpenAIRibbon_Load);
            this.openAITab.ResumeLayout(false);
            this.openAITab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab openAITab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton respondButton;
    }

    partial class ThisRibbonCollection
    {
        internal OpenAIRibbon OpenAIRibbon
        {
            get { return this.GetRibbon<OpenAIRibbon>(); }
        }
    }
}
