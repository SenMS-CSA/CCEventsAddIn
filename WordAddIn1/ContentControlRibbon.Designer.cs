using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    partial class ContentControlRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ContentControlRibbon()
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
            this.btnAddContentControl = this.Factory.CreateRibbonButton();
            this.btnAddEditableRegion = this.Factory.CreateRibbonButton();
            this.btnDeleteEditors = this.Factory.CreateRibbonButton();
            this.btnAddStructuredFigure = this.Factory.CreateRibbonButton();
            this.btnCreateNestedControls = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Content Controls";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnAddContentControl);
            this.group1.Items.Add(this.btnAddEditableRegion);
            this.group1.Items.Add(this.btnDeleteEditors);
            this.group1.Items.Add(this.btnAddStructuredFigure);
            this.group1.Items.Add(this.btnCreateNestedControls);
            this.group1.Label = "Insert Controls";
            this.group1.Name = "group1";
            // 
            // btnAddContentControl
            // 
            this.btnAddContentControl.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddContentControl.Label = "Add Content Control";
            this.btnAddContentControl.Name = "btnAddContentControl";
            this.btnAddContentControl.OfficeImageId = "MasterDocumentInsertSubdocument";
            this.btnAddContentControl.ShowImage = true;
            this.btnAddContentControl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddContentControl_Click);
            // 
            // btnAddEditableRegion
            // 
            this.btnAddEditableRegion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddEditableRegion.Label = "Add Editable Region";
            this.btnAddEditableRegion.Name = "btnAddEditableRegion";
            this.btnAddEditableRegion.OfficeImageId = "EditCitationButton";
            this.btnAddEditableRegion.ShowImage = true;
            this.btnAddEditableRegion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddEditableRegion_Click);
            // 
            // btnDeleteEditors
            // 
            this.btnDeleteEditors.Label = "Delete Editors";
            this.btnDeleteEditors.Name = "btnDeleteEditors";
            this.btnDeleteEditors.OfficeImageId = "Delete";
            this.btnDeleteEditors.ShowImage = true;
            this.btnDeleteEditors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDeleteEditors_Click);
            // 
            // btnAddStructuredFigure
            // 
            this.btnAddStructuredFigure.Label = "Add Structured Table";
            this.btnAddStructuredFigure.Name = "btnAddStructuredFigure";
            this.btnAddStructuredFigure.OfficeImageId = "CellsInsertDialog";
            this.btnAddStructuredFigure.ShowImage = true;
            this.btnAddStructuredFigure.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddStructuredFigure_Click);
            // 
            // btnCreateNestedControls
            // 
            this.btnCreateNestedControls.Label = "Create Nested Controls";
            this.btnCreateNestedControls.Name = "btnCreateNestedControls";
            // 
            // ContentControlRibbon
            // 
            this.Name = "ContentControlRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ContentControlRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddContentControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddEditableRegion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteEditors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddStructuredFigure;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateNestedControls;
    }
}
