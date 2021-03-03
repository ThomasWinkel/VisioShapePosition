
namespace ShapePosition
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tabTools = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnPositionExcel = this.Factory.CreateRibbonButton();
            this.btnDuplicateInExcel = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnReplace = this.Factory.CreateRibbonButton();
            this.tabTools.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabTools
            // 
            this.tabTools.Groups.Add(this.group1);
            this.tabTools.Groups.Add(this.group2);
            this.tabTools.Label = "Tools";
            this.tabTools.Name = "tabTools";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnPositionExcel);
            this.group1.Items.Add(this.btnDuplicateInExcel);
            this.group1.Label = "Position";
            this.group1.Name = "group1";
            // 
            // btnPositionExcel
            // 
            this.btnPositionExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPositionExcel.Label = "Modify in Excel";
            this.btnPositionExcel.Name = "btnPositionExcel";
            this.btnPositionExcel.OfficeImageId = "ExcelSpreadsheetInsert";
            this.btnPositionExcel.ShowImage = true;
            this.btnPositionExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPositionExcel_Click);
            // 
            // btnDuplicateInExcel
            // 
            this.btnDuplicateInExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDuplicateInExcel.Label = "Duplicate with Excel";
            this.btnDuplicateInExcel.Name = "btnDuplicateInExcel";
            this.btnDuplicateInExcel.OfficeImageId = "ExcelSpreadsheetInsert";
            this.btnDuplicateInExcel.ShowImage = true;
            this.btnDuplicateInExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDuplicateInExcel_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnReplace);
            this.group2.Label = "Tools";
            this.group2.Name = "group2";
            // 
            // btnReplace
            // 
            this.btnReplace.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReplace.Label = "Replace";
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.OfficeImageId = "ReplaceDialog";
            this.btnReplace.ShowImage = true;
            this.btnReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplace_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Visio.Drawing";
            this.Tabs.Add(this.tabTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tabTools.ResumeLayout(false);
            this.tabTools.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPositionExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDuplicateInExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplace;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
