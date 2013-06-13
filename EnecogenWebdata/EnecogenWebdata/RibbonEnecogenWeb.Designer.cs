namespace EnecogenWebdata
{
    partial class RibbonEnecogenWeb : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonEnecogenWeb()
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonEnecogenWeb));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.EnecogenWebGroup = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.DataType = this.Factory.CreateRibbonDropDown();
            this.dropDownNumberOfDays = this.Factory.CreateRibbonDropDown();
            this.dropDownRefreshInterval = this.Factory.CreateRibbonDropDown();
            this.editBoxSheetName = this.Factory.CreateRibbonEditBox();
            this.buttonStart = this.Factory.CreateRibbonButton();
            this.buttonStop = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.EnecogenWebGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.EnecogenWebGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // EnecogenWebGroup
            // 
            this.EnecogenWebGroup.Items.Add(this.separator1);
            this.EnecogenWebGroup.Items.Add(this.DataType);
            this.EnecogenWebGroup.Items.Add(this.dropDownNumberOfDays);
            this.EnecogenWebGroup.Items.Add(this.dropDownRefreshInterval);
            this.EnecogenWebGroup.Items.Add(this.editBoxSheetName);
            this.EnecogenWebGroup.Items.Add(this.buttonStart);
            this.EnecogenWebGroup.Items.Add(this.buttonStop);
            this.EnecogenWebGroup.Label = "Enecogen Web Data";
            this.EnecogenWebGroup.Name = "EnecogenWebGroup";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // DataType
            // 
            ribbonDropDownItemImpl1.Label = "Balance Delta with IGCC ";
            ribbonDropDownItemImpl1.Tag = "balancedeltaIGCC";
            this.DataType.Items.Add(ribbonDropDownItemImpl1);
            this.DataType.Label = "Data Type Parameter";
            this.DataType.Name = "DataType";
            this.DataType.ScreenTip = "Which type of Data from Tennet do you want";
            this.DataType.SizeString = "xxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            this.DataType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DataTypeParameter_Changes);
            // 
            // dropDownNumberOfDays
            // 
            ribbonDropDownItemImpl2.Label = "0";
            ribbonDropDownItemImpl2.Tag = "0";
            ribbonDropDownItemImpl3.Label = "1";
            ribbonDropDownItemImpl3.Tag = "1";
            ribbonDropDownItemImpl4.Label = "2";
            ribbonDropDownItemImpl4.Tag = "2";
            ribbonDropDownItemImpl5.Label = "3";
            ribbonDropDownItemImpl5.Tag = "3";
            ribbonDropDownItemImpl6.Label = "4";
            ribbonDropDownItemImpl6.Tag = "4";
            ribbonDropDownItemImpl7.Label = "5";
            ribbonDropDownItemImpl7.Tag = "5";
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl2);
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl3);
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl4);
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl5);
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl6);
            this.dropDownNumberOfDays.Items.Add(ribbonDropDownItemImpl7);
            this.dropDownNumberOfDays.Label = "Number of days";
            this.dropDownNumberOfDays.Name = "dropDownNumberOfDays";
            this.dropDownNumberOfDays.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDownNumberOfDays_SelectionChanged);
            // 
            // dropDownRefreshInterval
            // 
            ribbonDropDownItemImpl8.Label = "1 Minute";
            ribbonDropDownItemImpl8.Tag = "1";
            ribbonDropDownItemImpl9.Label = "2 Minutes";
            ribbonDropDownItemImpl9.Tag = "2";
            ribbonDropDownItemImpl10.Label = "3 Minutes";
            ribbonDropDownItemImpl10.Tag = "3";
            ribbonDropDownItemImpl11.Label = "4 Minutes";
            ribbonDropDownItemImpl11.Tag = "4";
            ribbonDropDownItemImpl12.Label = "5 Minutes";
            ribbonDropDownItemImpl12.Tag = "5";
            this.dropDownRefreshInterval.Items.Add(ribbonDropDownItemImpl8);
            this.dropDownRefreshInterval.Items.Add(ribbonDropDownItemImpl9);
            this.dropDownRefreshInterval.Items.Add(ribbonDropDownItemImpl10);
            this.dropDownRefreshInterval.Items.Add(ribbonDropDownItemImpl11);
            this.dropDownRefreshInterval.Items.Add(ribbonDropDownItemImpl12);
            this.dropDownRefreshInterval.Label = "Refresh Interval";
            this.dropDownRefreshInterval.Name = "dropDownRefreshInterval";
            // 
            // editBoxSheetName
            // 
            this.editBoxSheetName.Enabled = false;
            this.editBoxSheetName.Label = "Sheet Name";
            this.editBoxSheetName.Name = "editBoxSheetName";
            this.editBoxSheetName.SizeString = "xxxxxxxxxxxxxxxxxxxxxxxxxxxx";
            this.editBoxSheetName.Text = null;
            this.editBoxSheetName.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBoxSheetName_TextChanged);
            // 
            // buttonStart
            // 
            this.buttonStart.Image = global::EnecogenWebdata.Properties.Resources.go;
            this.buttonStart.Label = "Start Collecting";
            this.buttonStart.Name = "buttonStart";
            this.buttonStart.ShowImage = true;
            this.buttonStart.SuperTip = "Start collection of data";
            this.buttonStart.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStart_Click);
            // 
            // buttonStop
            // 
            this.buttonStop.Image = ((System.Drawing.Image)(resources.GetObject("buttonStop.Image")));
            this.buttonStop.Label = "Stop Collecting";
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.ShowImage = true;
            this.buttonStop.SuperTip = "Stop collection of data";
            this.buttonStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonStop_Click);
            // 
            // RibbonEnecogenWeb
            // 
            this.Name = "RibbonEnecogenWeb";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonEnecogenWeb_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.EnecogenWebGroup.ResumeLayout(false);
            this.EnecogenWebGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup EnecogenWebGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown DataType;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownNumberOfDays;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownRefreshInterval;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxSheetName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonStop;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonEnecogenWeb RibbonEnecogenWeb
        {
            get { return this.GetRibbon<RibbonEnecogenWeb>(); }
        }
    }
}
