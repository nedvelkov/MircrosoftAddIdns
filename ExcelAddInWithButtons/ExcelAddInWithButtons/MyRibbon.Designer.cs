
namespace ExcelAddInWithButtons
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
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
            this.customGroup = this.Factory.CreateRibbonGroup();
            this.toggleBold = this.Factory.CreateRibbonToggleButton();
            this.drawTable = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.customGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.customGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // customGroup
            // 
            this.customGroup.Items.Add(this.toggleBold);
            this.customGroup.Items.Add(this.drawTable);
            this.customGroup.Label = "Custom group";
            this.customGroup.Name = "customGroup";
            // 
            // toggleBold
            // 
            this.toggleBold.Label = "Toggle bold";
            this.toggleBold.Name = "toggleBold";
            this.toggleBold.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleBold_Click);
            // 
            // drawTable
            // 
            this.drawTable.Label = "Draw table";
            this.drawTable.Name = "drawTable";
            this.drawTable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.drawTable_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.customGroup.ResumeLayout(false);
            this.customGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup customGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleBold;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton drawTable;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
