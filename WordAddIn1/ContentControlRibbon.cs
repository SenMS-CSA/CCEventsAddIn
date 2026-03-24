using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class ContentControlRibbon
    {
        private void ContentControlRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.btnCreateNestedControls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateNestedControls_Click);
        }

        private void btnAddContentControl_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.AddContentControlWithRichText();
            Globals.ThisAddIn.ReproAddRemoveEditor();
        }

        private void btnAddEditableRegion_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddEditableRegionWithProtection();
        }

        private void btnDeleteEditors_Click(object sender, RibbonControlEventArgs e)
        {
            // Add your delete editors logic here
            Globals.ThisAddIn.DeleteControlsFromMasterControl();
        }

        private void btnAddStructuredFigure_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.AddStructuredFigureLayout();
        }

        private void btnCreateNestedControls_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CreateNestedControls();
        }
    }
}
