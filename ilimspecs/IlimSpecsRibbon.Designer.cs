namespace ilimspecs
{
    partial class IlimSpecsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IlimSpecsRibbon()
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
            this.ProductionUnitGroup = this.Factory.CreateRibbonGroup();
            this.btnImportProductionUnit = this.Factory.CreateRibbonButton();
            this.btnExportProductionUnit = this.Factory.CreateRibbonButton();
            this.productGroup = this.Factory.CreateRibbonGroup();
            this.btnImportProduct = this.Factory.CreateRibbonButton();
            this.btnExprotProduct = this.Factory.CreateRibbonButton();
            this.varSpecGroup = this.Factory.CreateRibbonGroup();
            this.btnImportVarSpec = this.Factory.CreateRibbonButton();
            this.btnExportVarSpec = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ProductionUnitGroup.SuspendLayout();
            this.productGroup.SuspendLayout();
            this.varSpecGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.ProductionUnitGroup);
            this.tab1.Groups.Add(this.productGroup);
            this.tab1.Groups.Add(this.varSpecGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // ProductionUnitGroup
            // 
            this.ProductionUnitGroup.Items.Add(this.btnImportProductionUnit);
            this.ProductionUnitGroup.Items.Add(this.btnExportProductionUnit);
            this.ProductionUnitGroup.Label = "Оборудование";
            this.ProductionUnitGroup.Name = "ProductionUnitGroup";
            // 
            // btnImportProductionUnit
            // 
            this.btnImportProductionUnit.Label = "Импорт оборудования";
            this.btnImportProductionUnit.Name = "btnImportProductionUnit";
            this.btnImportProductionUnit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportProductionUnit_Click);
            // 
            // btnExportProductionUnit
            // 
            this.btnExportProductionUnit.Label = "Экспорт оборудования";
            this.btnExportProductionUnit.Name = "btnExportProductionUnit";
            this.btnExportProductionUnit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportProductionUnit_Click);
            // 
            // productGroup
            // 
            this.productGroup.Items.Add(this.btnImportProduct);
            this.productGroup.Items.Add(this.btnExprotProduct);
            this.productGroup.Label = "Продукты";
            this.productGroup.Name = "productGroup";
            // 
            // btnImportProduct
            // 
            this.btnImportProduct.Label = "Импорт продуктов";
            this.btnImportProduct.Name = "btnImportProduct";
            this.btnImportProduct.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportProduct_Click);
            // 
            // btnExprotProduct
            // 
            this.btnExprotProduct.Label = "Экспорт продуктов";
            this.btnExprotProduct.Name = "btnExprotProduct";
            this.btnExprotProduct.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExprotProduct_Click);
            // 
            // varSpecGroup
            // 
            this.varSpecGroup.Items.Add(this.btnImportVarSpec);
            this.varSpecGroup.Items.Add(this.btnExportVarSpec);
            this.varSpecGroup.Label = "Спецификации";
            this.varSpecGroup.Name = "varSpecGroup";
            // 
            // btnImportVarSpec
            // 
            this.btnImportVarSpec.Label = "Импорт спецификаций";
            this.btnImportVarSpec.Name = "btnImportVarSpec";
            this.btnImportVarSpec.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportVarSpec_Click);
            // 
            // btnExportVarSpec
            // 
            this.btnExportVarSpec.Label = "Экспорт спецификаций";
            this.btnExportVarSpec.Name = "btnExportVarSpec";
            this.btnExportVarSpec.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportVarSpec_Click);
            // 
            // IlimSpecsRibbon
            // 
            this.Name = "IlimSpecsRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IlimSpecsRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ProductionUnitGroup.ResumeLayout(false);
            this.ProductionUnitGroup.PerformLayout();
            this.productGroup.ResumeLayout(false);
            this.productGroup.PerformLayout();
            this.varSpecGroup.ResumeLayout(false);
            this.varSpecGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ProductionUnitGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportProductionUnit;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportProductionUnit;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup productGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportProduct;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExprotProduct;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup varSpecGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportVarSpec;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportVarSpec;
    }

    partial class ThisRibbonCollection
    {
        internal IlimSpecsRibbon IlimSpecsRibbon
        {
            get { return this.GetRibbon<IlimSpecsRibbon>(); }
        }
    }
}
