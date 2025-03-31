namespace ArasExport_AddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.Login = this.Factory.CreateRibbonGroup();
            this.Btn_Login = this.Factory.CreateRibbonButton();
            this.Update = this.Factory.CreateRibbonGroup();
            this.Btn_Validate = this.Factory.CreateRibbonButton();
            this.Btn_Resolve = this.Factory.CreateRibbonButton();
            this.Export = this.Factory.CreateRibbonGroup();
            this.Btn_Export = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Login.SuspendLayout();
            this.Update.SuspendLayout();
            this.Export.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Login);
            this.tab1.Groups.Add(this.Update);
            this.tab1.Groups.Add(this.Export);
            this.tab1.Label = "Aras Utility";
            this.tab1.Name = "tab1";
            // 
            // Login
            // 
            this.Login.Items.Add(this.Btn_Login);
            this.Login.Label = "Login Tools";
            this.Login.Name = "Login";
            // 
            // Btn_Login
            // 
            this.Btn_Login.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_Login.Image = global::ArasExport_AddIn.Properties.Resources.login;
            this.Btn_Login.Label = "Login";
            this.Btn_Login.Name = "Btn_Login";
            this.Btn_Login.ShowImage = true;
            this.Btn_Login.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Login_Click);
            // 
            // Update
            // 
            this.Update.Items.Add(this.Btn_Validate);
            this.Update.Items.Add(this.Btn_Resolve);
            this.Update.Label = "Update Tools";
            this.Update.Name = "Update";
            // 
            // Btn_Validate
            // 
            this.Btn_Validate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_Validate.Enabled = false;
            this.Btn_Validate.Image = global::ArasExport_AddIn.Properties.Resources.update_elements;
            this.Btn_Validate.Label = "Update Excel";
            this.Btn_Validate.Name = "Btn_Validate";
            this.Btn_Validate.ShowImage = true;
            this.Btn_Validate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Validate_Click);
            // 
            // Btn_Resolve
            // 
            this.Btn_Resolve.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_Resolve.Enabled = false;
            this.Btn_Resolve.Image = global::ArasExport_AddIn.Properties.Resources.add_to_pkg;
            this.Btn_Resolve.Label = "Update Package";
            this.Btn_Resolve.Name = "Btn_Resolve";
            this.Btn_Resolve.ShowImage = true;
            this.Btn_Resolve.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Resolve_Click);
            // 
            // Export
            // 
            this.Export.Items.Add(this.Btn_Export);
            this.Export.Label = "Export Tool";
            this.Export.Name = "Export";
            // 
            // Btn_Export
            // 
            this.Btn_Export.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Btn_Export.Enabled = false;
            this.Btn_Export.Image = global::ArasExport_AddIn.Properties.Resources.export;
            this.Btn_Export.Label = "Export";
            this.Btn_Export.Name = "Btn_Export";
            this.Btn_Export.ShowImage = true;
            this.Btn_Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Btn_Export_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Login.ResumeLayout(false);
            this.Login.PerformLayout();
            this.Update.ResumeLayout(false);
            this.Update.PerformLayout();
            this.Export.ResumeLayout(false);
            this.Export.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Login;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Update;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_Login;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_Validate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_Resolve;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Btn_Export;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
