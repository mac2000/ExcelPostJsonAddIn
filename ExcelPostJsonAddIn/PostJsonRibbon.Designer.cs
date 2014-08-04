namespace ExcelPostJsonAddIn
{
    partial class PostJsonRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PostJsonRibbon()
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
            this.groupPostJson = this.Factory.CreateRibbonGroup();
            this.editBoxUrl = this.Factory.CreateRibbonEditBox();
            this.editBoxUser = this.Factory.CreateRibbonEditBox();
            this.buttonSubmit = this.Factory.CreateRibbonButton();
            this.editBoxPass = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.groupPostJson.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupPostJson);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // groupPostJson
            // 
            this.groupPostJson.Items.Add(this.editBoxUrl);
            this.groupPostJson.Items.Add(this.editBoxUser);
            this.groupPostJson.Items.Add(this.editBoxPass);
            this.groupPostJson.Items.Add(this.buttonSubmit);
            this.groupPostJson.Label = "Post JSON";
            this.groupPostJson.Name = "groupPostJson";
            // 
            // editBoxUrl
            // 
            this.editBoxUrl.Label = "URL";
            this.editBoxUrl.Name = "editBoxUrl";
            this.editBoxUrl.ScreenTip = "URL to post JSON data to";
            // 
            // editBoxUser
            // 
            this.editBoxUser.Label = "User";
            this.editBoxUser.Name = "editBoxUser";
            this.editBoxUser.ScreenTip = "HTTP Auth Username (Optional)";
            // 
            // buttonSubmit
            // 
            this.buttonSubmit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonSubmit.Enabled = false;
            this.buttonSubmit.Image = global::ExcelPostJsonAddIn.Properties.Resources.icon64;
            this.buttonSubmit.Label = "Submit";
            this.buttonSubmit.Name = "buttonSubmit";
            this.buttonSubmit.ShowImage = true;
            this.buttonSubmit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSubmit_Click);
            // 
            // editBoxPass
            // 
            this.editBoxPass.Label = "Pass";
            this.editBoxPass.Name = "editBoxPass";
            this.editBoxPass.ScreenTip = "HTTP Auth Password (Optional)";
            // 
            // PostJsonRibbon
            // 
            this.Name = "PostJsonRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PostJsonRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupPostJson.ResumeLayout(false);
            this.groupPostJson.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupPostJson;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSubmit;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxUrl;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxUser;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxPass;
    }

    partial class ThisRibbonCollection
    {
        internal PostJsonRibbon PostJsonRibbon
        {
            get { return this.GetRibbon<PostJsonRibbon>(); }
        }
    }
}
