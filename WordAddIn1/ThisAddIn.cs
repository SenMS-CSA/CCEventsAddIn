using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            RegisterDocumentEventHandlers();
            // StartContentMonitoring(); // Disable temporarily
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            StopContentMonitoring();
            UnregisterDocumentEventHandlers();
        }

        #region Event Handler Registration

        private void RegisterDocumentEventHandlers()
        {
            // Document events
            this.Application.DocumentOpen += Application_DocumentOpen;
            this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
            this.Application.DocumentChange += Application_DocumentChange;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // Content control events - these fire when content controls are modified
            //((Word.ApplicationEvents4_Event)this.Application).ContentControlOnEnter += Application_ContentControlOnEnter;
            //((Word.ApplicationEvents4_Event)this.Application).ContentControlOnExit += Application_ContentControlOnExit;
            //((Word.ApplicationEvents4_Event)this.Application).ContentControlBeforeDelete += Application_ContentControlBeforeDelete;

            TraceLog("Event handlers registered");
        }

        private void UnregisterDocumentEventHandlers()
        {
            try
            {
                this.Application.DocumentOpen -= Application_DocumentOpen;
                this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
                this.Application.DocumentBeforeClose -= Application_DocumentBeforeClose;
                this.Application.DocumentChange -= Application_DocumentChange;
                this.Application.WindowSelectionChange -= Application_WindowSelectionChange;

                TraceLog("Event handlers unregistered");
            }
            catch (Exception ex)
            {
                TraceLog("Error unregistering events: " + ex.Message);
            }
        }

        #endregion

        #region Document Event Handlers

        private void Application_DocumentOpen(Word.Document Doc)
        {
            TraceLog($"DocumentOpen: {Doc.Name}");
        }

        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            TraceLog($"DocumentBeforeSave: {Doc.Name}, SaveAsUI={SaveAsUI}");
        }

        private void Application_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            TraceLog($"DocumentBeforeClose: {Doc.Name}");
        }

        private void Application_DocumentChange()
        {
            try
            {
                Word.Document doc = this.Application.ActiveDocument;
                string docName = doc != null ? doc.Name : "No active document";
                TraceLog($"DocumentChange: {docName}");
            }
            catch
            {
                TraceLog("DocumentChange: Unable to get active document");
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                string info = $"SelectionChange: Start={Sel.Start}, End={Sel.End}";
                if (Sel.Range.ContentControls.Count > 0)
                {
                    Word.ContentControl cc = Sel.Range.ContentControls[1];
                    info += $", ContentControl: {cc.Tag}";
                }
                TraceLog(info);
            }
            catch (Exception ex)
            {
                TraceLog($"SelectionChange error: {ex.Message}");
            }
        }

        #endregion

        #region Content Control Event Handlers

        private void Application_ContentControlOnEnter(Word.ContentControl ContentControl)
        {
            TraceLog($"ContentControlOnEnter: Tag={ContentControl.Tag}, Title={ContentControl.Title}");
        }

        private void Application_ContentControlOnExit(Word.ContentControl ContentControl, ref bool Cancel)
        {
            TraceLog($"ContentControlOnExit: Tag={ContentControl.Tag}, Title={ContentControl.Title}");
        }

        private void Application_ContentControlBeforeContentUpdate(Word.ContentControl ContentControl, ref string Content)
        {
            TraceLog($"ContentControlBeforeContentUpdate: Tag={ContentControl.Tag}");
        }

        private void Application_ContentControlAfterAdd(Word.ContentControl ContentControl, bool InUndoRedo)
        {
            TraceLog($"ContentControlAfterAdd: Tag={ContentControl.Tag}, InUndoRedo={InUndoRedo}");
        }

        private void Application_ContentControlBeforeDelete(Word.ContentControl ContentControl, bool InUndoRedo)
        {
            TraceLog($"ContentControlBeforeDelete: Tag={ContentControl.Tag}, InUndoRedo={InUndoRedo}");
        }

        #endregion

        #region Tracing Helper

        private void TraceLog(string message)
        {
            string logMessage = "[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + message;
            Debug.WriteLine(logMessage);
        }

        #endregion

        /// <summary>
        /// Adds a content control with a rich text control inside and demonstrates range editing
        /// </summary>
        public void AddContentControlWithRichText()
        {
            Word.Document doc = this.Application.ActiveDocument;

            // Get the current selection or create a range
            Word.Range range = this.Application.Selection.Range;

            // Create the outer master content control that wraps everything
            Word.ContentControl outerMasterControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                range);
            outerMasterControl.Title = "Outer Master Control";
            outerMasterControl.Tag = "OuterMasterControl";

            // Get the range inside the outer master control
            Word.Range outerRange = outerMasterControl.Range;

            // Insert a 1x2 table for side-by-side figure controls
            Word.Table table = doc.Tables.Add(outerRange, 1, 2);
            table.Borders.Enable = 0; // No borders on outer table
            table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100;

            // Set column widths to be equal
            table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[1].PreferredWidth = 50;
            table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[2].PreferredWidth = 50;

            // Create Figure 1 bundle in first cell
            CreateFigureBundle(doc, table.Cell(1, 1), "Figure 1. Test image 1", "FigureControl1", "ImageControl1", "FooterControl1");

            // Create Figure 2 bundle in second cell
            CreateFigureBundle(doc, table.Cell(1, 2), "Figure 2. Test image 2", "FigureControl2", "ImageControl2", "FooterControl2");

            // Move cursor after the outer master control
            Word.Range afterRange = outerMasterControl.Range;
            afterRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterRange.Select();
        }


        private void CreateFigureBundle(Word.Document doc, Word.Cell cell, string title, string figureTag, string imageTag, string footerTag)
        {
            // Get the cell range
            Word.Range cellRange = cell.Range;

            // Create the figure content control that wraps everything for this figure
            Word.ContentControl figureControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                cellRange);
            figureControl.Title = title;
            figureControl.Tag = figureTag;

            // Get the range inside the figure control
            Word.Range figureRange = figureControl.Range;

            // Add header text
            figureRange.Text = title + "\n";
            figureRange.Font.Bold = 1;
            figureRange.Font.Size = 11;
            figureRange.Font.Color = Word.WdColor.wdColorBlack;
            figureRange.ParagraphFormat.SpaceAfter = 6;

            // Move to end to add image placeholder
            Word.Range imageRange = figureControl.Range;
            imageRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            imageRange.InsertParagraph();
            imageRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            imageRange.Text = "[Image or content placeholder]";
            imageRange.Font.Bold = 0;
            imageRange.Font.Size = 10;
            imageRange.Font.Color = Word.WdColor.wdColorGray50;
            imageRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            imageRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

            // Create image content control
            Word.ContentControl imageControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                imageRange);
            imageControl.Title = "Image Placeholder";
            imageControl.Tag = imageTag;

            // Set minimum height for image area by adding line breaks
            Word.Range imageInnerRange = imageControl.Range;
            imageInnerRange.Text = "\n\n\n\n[Image or content placeholder]\n\n\n\n";
            imageInnerRange.Font.Bold = 0;
            imageInnerRange.Font.Size = 10;
            imageInnerRange.Font.Color = Word.WdColor.wdColorGray50;
            imageInnerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            imageInnerRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

            // Add footer text after the image
            Word.Range footerRange = figureControl.Range;
            footerRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            footerRange.MoveEnd(Word.WdUnits.wdCharacter, -1); // Move back before the cell marker
            footerRange.InsertAfter("\r"); // Insert paragraph mark using InsertAfter instead
            footerRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            footerRange.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.\nSource: XYZ Research";

            // Make footer text non-editable by adding it to a locked content control
            Word.ContentControl footerControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                footerRange);
            footerControl.Title = "Footer (Read Only)";
            footerControl.Tag = footerTag;
            footerControl.LockContents = true; // Lock the contents to prevent editing
        }


        /// <summary>
        /// Creates a complete figure bundle with header, image placeholder, and footer
        /// </summary>
        private void CreateFigureBundle(Word.Document doc, Word.Range targetRange, string title, string figureTag, string imageTag, string footerTag)
        {
            // Create the figure content control that wraps everything for this figure
            Word.ContentControl figureControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                targetRange);
            figureControl.Title = title;
            figureControl.Tag = figureTag;

            // Get the range inside the figure control
            Word.Range figureRange = figureControl.Range;

            // Add header text
            figureRange.Text = title + "\n";
            figureRange.Font.Bold = 1;
            figureRange.Font.Size = 11;
            figureRange.Font.Color = Word.WdColor.wdColorBlack;
            figureRange.ParagraphFormat.SpaceAfter = 6;

            // Move to end to add image placeholder
            Word.Range imageRange = figureControl.Range;
            imageRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            imageRange.InsertParagraph();
            imageRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            imageRange.Text = "[Image or content placeholder]";
            imageRange.Font.Bold = 0;
            imageRange.Font.Size = 10;
            imageRange.Font.Color = Word.WdColor.wdColorGray50;
            imageRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            imageRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

            // Create image content control
            Word.ContentControl imageControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                imageRange);
            imageControl.Title = "Image Placeholder";
            imageControl.Tag = imageTag;

            // Set minimum height for image area by adding line breaks
            Word.Range imageInnerRange = imageControl.Range;
            imageInnerRange.Text = "\n\n\n\n[Image or content placeholder]\n\n\n\n";
            imageInnerRange.Font.Bold = 0;
            imageInnerRange.Font.Size = 10;
            imageInnerRange.Font.Color = Word.WdColor.wdColorGray50;
            imageInnerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            imageInnerRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

            // Add footer text after the image
            Word.Range footerRange = figureControl.Range;
            footerRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            footerRange.MoveEnd(Word.WdUnits.wdCharacter, -1); // Move back before the cell marker
            footerRange.InsertAfter("\r"); // Insert paragraph mark using InsertAfter instead
            footerRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            footerRange.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.\nSource: XYZ Research";

            // Make footer text non-editable by adding it to a locked content control
            Word.ContentControl footerControl = doc.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                footerRange);
            footerControl.Title = "Footer (Read Only)";
            footerControl.Tag = footerTag;
            footerControl.LockContents = true; // Lock the contents to prevent editing
        }

        /// <summary>
        /// Demonstrates using Range.Editors for document protection scenarios
        /// </summary>
        public void AddEditableRegionWithProtection()
        {
            Word.Document doc = this.Application.ActiveDocument;
            Word.Range range = this.Application.Selection.Range;

            // Add content to the range
            range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
            //range.Editors.Item(0).Delete();
            range.Text = "This region is editable even when document is protected";

            // Add the range as an editable region
            Word.Editor editor = range.Editors.Add(Word.WdEditorType.wdEditorEveryone);


            //range.Editors.Item(0).Delete();

            // Format the editable region
            range.Font.Color = Word.WdColor.wdColorGreen;
            range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            // Create a 1x1 table to contain the figure
            Word.Range outerRange = range;

            // Insert a 1x2 table for side-by-side figure controls
            Word.Table table = doc.Tables.Add(range, 1, 2);
            table.Borders.Enable = 0; // No borders on outer table
            table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100;

            // Set column widths to be equal
            table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[1].PreferredWidth = 50;
            table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.Columns[2].PreferredWidth = 50;
            CreateFigureBundle(doc, table.Cell(1, 1), "Editable Figure. Test image1", "EditableFigureControl", "EditableImageControl", "EditableFooterControl");

            CreateFigureBundle(doc, table.Cell(1, 2), "Editable Figure. Test image2", "EditableFigureControl", "EditableImageControl", "EditableFooterControl");

            //CreateFigureBundle(doc, range, "Editable Figure. Test image", "EditableFigureControl", "EditableImageControl", "EditableFooterControl");

            // You can now protect the document, but this range remains editable
            // doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, false, "", false, false);
        }

        /// <summary>
        /// Deletes all editors from the master content control
        /// </summary>
        public void DeleteEditorsFromMasterControl()
        {
            Word.Document doc = this.Application.ActiveDocument;
            
            // Find the outer master control by tag
            Word.ContentControl masterControl = null;
            foreach (Word.ContentControl cc in doc.ContentControls)
            {
                if (cc.Tag == "OuterMasterControl")
                {
                    masterControl = cc;
                    break;
                }
            }
            
            if (masterControl == null)
            {
                System.Windows.Forms.MessageBox.Show("Master content control not found.", "Delete Editors", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            
            Word.Range masterRange = masterControl.Range;
            
            // Delete editors by looping backwards (to avoid index shifting issues)
            int editorCount = masterRange.Editors.Count;
            for (int i = editorCount; i >= 1; i--)
            {
                masterRange.Editors.Item(i).Delete();
            }
            
            System.Windows.Forms.MessageBox.Show($"Deleted {editorCount} editor(s) from the master control.", "Delete Editors",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        /// <summary>
        /// Deletes all content controls and tables from the master content control
        /// </summary>
        public void DeleteControlsFromMasterControl()
        {
            Word.Document doc = this.Application.ActiveDocument;
            
            // Find the outer master control by tag
            Word.ContentControl masterControl = null;
            foreach (Word.ContentControl cc in doc.ContentControls)
            {
                if (cc.Tag == "boxtype") //charts##3294615616 OuterMasterControl
                {
                    masterControl = cc;
                    break;
                }
            }
            
            if (masterControl == null)
            {
                System.Windows.Forms.MessageBox.Show("Master content control not found.", "Delete Controls", 
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            
            Word.Range masterRange = masterControl.Range;

            // Delete content controls by looping backwards -Temporarily disabled to preserve figure controls for testing
            int controlCount = masterRange.ContentControls.Count;
            for (int i = controlCount; i >= 1; i--)
            {
                //Range.Editors.Item(i).Delete()
                // masterRange.Editors.Item(i).Delete();
                //    masterRange.ContentControls[i].Delete(true); // commented out to preserve figure controls for testing
            }

            // Delete tables by looping backwards
            int tableCount = masterRange.Tables.Count;
            for (int i = tableCount; i >= 1; i--)
            {
                masterRange.Tables[i].Delete();

            }
            
          //  System.Windows.Forms.MessageBox.Show($"Deleted {controlCount} content control(s) and {tableCount} table(s) from the master control.", "Delete Controls",
              //  System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        private Microsoft.Office.Tools.Ribbon.RibbonButton btnDeleteEditors;
        private void btnDeleteEditors_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DeleteEditorsFromMasterControl();
        }

        private System.Windows.Forms.Timer _changeTimer;
        private string _lastContent;
        private bool _isProcessing = false;

        private void StartContentMonitoring()
        {
            _changeTimer = new System.Windows.Forms.Timer();
            _changeTimer.Interval = 1000; // Increase to 1 second
            _changeTimer.Tick += CheckForContentChanges;
            _changeTimer.Start();
        }

        private void StopContentMonitoring()
        {
            if (_changeTimer != null)
            {
                _changeTimer.Stop();
                _changeTimer.Dispose();
                _changeTimer = null;
            }
        }

        private void CheckForContentChanges(object sender, EventArgs e)
        {
            // Prevent re-entrancy
            if (_isProcessing) return;
            
            _isProcessing = true;
            _changeTimer.Stop(); // Pause timer during check
            
            try
            {
                // Check if Word is ready
                if (this.Application == null) return;
                
                Word.Document doc = null;
                try
                {
                    doc = this.Application.ActiveDocument;
                }
                catch
                {
                    return; // Word is busy
                }
                
                if (doc == null) return;
                
                // Use character count instead of full text (much lighter)
                int currentCharCount = doc.Characters.Count;
                string currentHash = currentCharCount.ToString();
                
                if (_lastContent != null && currentHash != _lastContent)
                {
                    TraceLog($"Content changed: CharCount={currentCharCount}");
                }
                _lastContent = currentHash;
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                TraceLog($"COM error (Word busy): {comEx.Message}");
            }
            catch (Exception ex)
            {
                TraceLog($"CheckForContentChanges error: {ex.Message}");
            }
            finally
            {
                _isProcessing = false;
                if (_changeTimer != null)
                {
                    _changeTimer.Start(); // Resume timer
                }
            }
        }

        /// <summary>
        /// Creates a structured figure layout with editable regions and content controls
        /// </summary>
        public void AddStructuredFigureLayout()
        {
            Word.Document doc = this.Application.ActiveDocument;
            Word.Range range = this.Application.Selection.Range;

            // Store protection state and unprotect if needed
            Word.WdProtectionType originalProtection = doc.ProtectionType;
            if (originalProtection != Word.WdProtectionType.wdNoProtection)
            {
                doc.Unprotect(Password: "test");
            }

            try
            {
                // Step 1: Insert the table FIRST at the current selection (before creating content controls)
                Word.Table table = doc.Tables.Add(range, 4, 2);
                table.Borders.Enable = 1;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;

                // Set column widths to be equal
                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 50;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 50;

                // Step 2: Populate cell content
                // Cell (1,1) - "Figure 1. Test image 1"
                Word.Range cell11Range = table.Cell(1, 1).Range;
                cell11Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell11Range.Text = "Figure 1. Test image 1";
                cell11Range.Font.Bold = 1;
                cell11Range.Font.Size = 11;

                // Cell (1,2) - "Figure 2. Test image 2"
                Word.Range cell12Range = table.Cell(1, 2).Range;
                cell12Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell12Range.Text = "Figure 2. Test image 2";
                cell12Range.Font.Bold = 1;
                cell12Range.Font.Size = 11;

                // Cell (2,1) - Image placeholder 1
                
                Word.Range cell21Range = table.Cell(2, 1).Range;
                cell21Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell21Range.Text = "[Image Placeholder 1]";
                cell21Range.Font.Bold = 0;
                cell21Range.Font.Size = 10;
                cell21Range.Font.Color = Word.WdColor.wdColorGray50;
                cell21Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cell21Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

                // Cell (2,2) - Image placeholder 2
                Word.Range cell22Range = table.Cell(2, 2).Range;
                cell22Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell22Range.Text = "[Image Placeholder 2]";
                cell22Range.Font.Bold = 0;
                cell22Range.Font.Size = 10;
                cell22Range.Font.Color = Word.WdColor.wdColorGray50;
                cell22Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cell22Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;

                // Cells (3,1) and (3,2) - Copyright text
                Word.Range cell31Range = table.Cell(3, 1).Range;
                cell31Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell31Range.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.";
                cell31Range.Font.Bold = 0;
                cell31Range.Font.Size = 8;

                Word.Range cell32Range = table.Cell(3, 2).Range;
                cell32Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell32Range.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.";
                cell32Range.Font.Bold = 0;
                cell32Range.Font.Size = 8;

                // Cells (4,1) and (4,2) - Source text
                Word.Range cell41Range = table.Cell(4, 1).Range;
                cell41Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell41Range.Text = "Source: XYZ Research";
                cell41Range.Font.Bold = 0;
                cell41Range.Font.Size = 8;

                Word.Range cell42Range = table.Cell(4, 2).Range;
                cell42Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell42Range.Text = "Source: XYZ Research";
                cell42Range.Font.Bold = 0;
                cell42Range.Font.Size = 8;

                // Step 3: Get the range that encompasses the entire table
                Word.Range tableFullRange = table.Range;

                // Step 4: First wrap the table in the innermost control (Figure Layout Control)
                Word.ContentControl figureLayoutControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText,
                    tableFullRange);
                figureLayoutControl.Title = "Figure Layout Control";
                figureLayoutControl.Tag = "FigureLayoutControl";
                figureLayoutControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                figureLayoutControl.Color = Word.WdColor.wdColorBlue;

              
                // figureLayoutControl.SetPlaceholderText(null, null, "This is the Figure Layout Control. It wraps the entire figure including header, image, and footer.");
                // Step 5: To properly nest, we need to expand the range to INCLUDE the content control boundaries
                // The Range property only returns the contents, not the control tags
                // We expand by 1 character on each side to include the control start/end markers
                int innerStart = figureLayoutControl.Range.Start;
                int innerEnd = figureLayoutControl.Range.End;
                Word.Range middleControlRange = doc.Range(innerStart, innerEnd);
                // Expand to include control boundaries
                middleControlRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                middleControlRange.MoveEnd(Word.WdUnits.wdCharacter, 1);

                // Insert text BEFORE the figure layout (will be inside Editable Region)
                Word.Range beforeFigureRange = doc.Range(middleControlRange.Start, middleControlRange.Start);
                beforeFigureRange.InsertBefore("This paragraph appears before the figure inside the Editable Region.\r\n");
                beforeFigureRange.Font.Bold = 0;
                beforeFigureRange.Font.Size = 11;

                // Refresh middleControlRange to include the new text
                middleControlRange.SetRange(beforeFigureRange.Start, middleControlRange.End);

                // Insert text AFTER the figure layout (will be inside Editable Region)
                Word.Range afterFigureRange = doc.Range(middleControlRange.End, middleControlRange.End);
              //  afterFigureRange.InsertAfter("This paragraph appears after the figure inside the Editable Region.");
                afterFigureRange.Font.Bold = 0;
                afterFigureRange.Font.Size = 11;

                // Refresh middleControlRange to include both texts
                middleControlRange.SetRange(middleControlRange.Start, afterFigureRange.End);

                // Step 6: Now create the Editable Region Control wrapping everything
                Word.ContentControl editableRegionControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText,
                    middleControlRange);
                editableRegionControl.Title = "Editable Region Control";
                editableRegionControl.Tag = "EditableRegion";
                editableRegionControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                editableRegionControl.Color = Word.WdColor.wdColorGreen;
             
                // NOTE: Do NOT add an editor to the entire editableRegionControl.Range
                // as that would make ALL content inside editable. Instead, we add editors
                // only to specific cells (rows 1-2) in Step 9 below.

                // Step 8: Wrap everything in the outermost control (Body Text Control)
                int middleStart = editableRegionControl.Range.Start;
                int middleEnd = editableRegionControl.Range.End;
                Word.Range outerControlRange = doc.Range(middleStart, middleEnd);
                // Expand to include control boundaries
                outerControlRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                outerControlRange.MoveEnd(Word.WdUnits.wdCharacter, 1);
              
                Word.ContentControl bodyTextControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText,
                    outerControlRange);
                bodyTextControl.Title = "Body Text Control";
                bodyTextControl.Tag = "BodyText";
                bodyTextControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
               
                //bodyTextControl.Color = Word.WdColor.wdColorRed;

                // Add text at the beginning of the Body Text Control (before the nested controls)
                Word.Range startTextRange = bodyTextControl.Range;
                startTextRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
               // startTextRange.InsertBefore("=== Body Text Control Start ===\r\n");
                startTextRange.Font.Bold = 1;
                startTextRange.Font.Color = Word.WdColor.wdColorRed;

                // Add text at the end of the Body Text Control (after the nested controls)
                Word.Range endTextRange = bodyTextControl.Range;
                endTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endTextRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                endTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
              //  endTextRange.InsertAfter("\r\n=== Body Text Control End ===");
                // Format the end text
                Word.Range endFormatRange = bodyTextControl.Range;
                endFormatRange.SetRange(endFormatRange.End - 30, endFormatRange.End);
                endFormatRange.Font.Bold = 1;
               // endFormatRange.Font.Color = Word.WdColor.wdColorRed;

                // Step 9: Add editable regions to rows 1 and 2 ONLY (title and image rows)
                // These will be editable when document protection is applied
                for (int row = 1; row <= 2; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = table.Cell(row, col).Range;
                        cellRange.MoveEnd(Word.WdUnits.wdCharacter, -1);
                        cellRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                    }
                }

                // Rows 3 and 4 (copyright and source) are NOT given editable regions - they stay locked

                // Move cursor after the control
                Word.Range afterRange = bodyTextControl.Range;
                afterRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                afterRange.Select();

                TraceLog("Structured figure layout created successfully");
            }
            finally
            {
                // Protect the document with AllowOnlyReading to enforce editable regions
                doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, Password: "test");
            }
        }

        /// <summary>
        /// Creates nested content controls with random text, a table, and editable regions
        /// </summary>
public void CreateNestedControls()
{
    Word.Document doc = this.Application.ActiveDocument;
    Word.Range range = this.Application.Selection.Range;

    // Store protection state and unprotect if needed
    Word.WdProtectionType originalProtection = doc.ProtectionType;
    if (originalProtection != Word.WdProtectionType.wdNoProtection)
    {
        doc.Unprotect(Password: "test");
    }

            try
            {
                // Step 1: Create the Master Control (outermost content control) FIRST
                range.Text = " "; // Placeholder text for master control
                Word.ContentControl masterControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText,
                    range);
                masterControl.Title = "Master Control";
                masterControl.Tag = "BodyText";
                masterControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                masterControl.Color = Word.WdColor.wdColorRed;

                // Step 2: Add editable paragraph text inside Master Control
                Word.Range masterInnerRange = masterControl.Range;
                masterInnerRange.Text = "Sample paragraph text before the nested control. This text is inside the Master Control.\r\n\r\n";

                // Make the Master Control paragraph text editable
                Word.Range masterTextRange = doc.Range(masterControl.Range.Start, masterControl.Range.End - 1);
                masterTextRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Step 3: Create the Inner Control SECOND (inside Master Control)
                Word.Range innerControlRange = doc.Range(masterControl.Range.End - 1, masterControl.Range.End - 1);
                innerControlRange.Text = " "; // Placeholder for inner control

                Word.ContentControl innerControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText,
                    innerControlRange);
                innerControl.Title = "Figure Layout Control";
                innerControl.Tag = "FigureLayout";
                innerControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                innerControl.Color = Word.WdColor.wdColorBlue;

                // Step 4: Add editable text inside Inner Control (before table)
                Word.Range innerTextRange = innerControl.Range;
                innerTextRange.Text = "This text is inside the Inner Control, before the table.\r\n\r\n";

                // Make the Inner Control text editable
                Word.Range innerEditableRange = doc.Range(innerControl.Range.Start, innerControl.Range.End - 1);
                innerEditableRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Step 5: Insert the Table THIRD (inside Inner Control)
                Word.Range tableInsertRange = doc.Range(innerControl.Range.End - 1, innerControl.Range.End - 1);

                Word.Table table = doc.Tables.Add(tableInsertRange, 4, 2);
                table.Borders.Enable = 1;
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.PreferredWidth = 100;

                // Set column widths to be equal
                table.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[1].PreferredWidth = 50;
                table.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                table.Columns[2].PreferredWidth = 50;

                // Step 6: Add content to cells
                // Row 1 - Figure titles (EDITABLE)
                Word.Range cell11Range = table.Cell(1, 1).Range;
                cell11Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell11Range.Text = "Figure 1. Test image 1";
                cell11Range.Font.Bold = 1;
                cell11Range.Font.Size = 11;
                cell11Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Word.Range cell12Range = table.Cell(1, 2).Range;
                cell12Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell12Range.Text = "Figure 2. Test image 2";
                cell12Range.Font.Bold = 1;
                cell12Range.Font.Size = 11;
                cell12Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Row 2 - Image placeholders (EDITABLE)
                Word.Range cell21Range = table.Cell(2, 1).Range;
                cell21Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell21Range.Text = "[Image Placeholder 1]";
                cell21Range.Font.Bold = 0;
                cell21Range.Font.Size = 10;
                cell21Range.Font.Color = Word.WdColor.wdColorGray50;
                cell21Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cell21Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                cell21Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Word.Range cell22Range = table.Cell(2, 2).Range;
                cell22Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell22Range.Text = "[Image Placeholder 2]";
                cell22Range.Font.Bold = 0;
                cell22Range.Font.Size = 10;
                cell22Range.Font.Color = Word.WdColor.wdColorGray50;
                cell22Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cell22Range.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                cell22Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Row 3 - Copyright text (NOT EDITABLE - no editors added)
                Word.Range cell31Range = table.Cell(3, 1).Range;
                cell31Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell31Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                //cell31Range.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.";

                //Below line causes the object has been deleted error in line 880, so it means the next range is getting deleted when assign a text to the range
                cell31Range.Editors.Item(1).Range.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.";
                cell31Range.Font.Bold = 0;
                cell31Range.Font.Size = 8;
                //  cell31Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Word.Range cell32Range = table.Cell(3, 2).Range;
                cell32Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell32Range.Text = "© 2025 XYZ Inc. No redistribution without group's written permission.";
                cell32Range.Font.Bold = 0;
                cell32Range.Font.Size = 8;
                //cell32Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Row 4 - Source text (NOT EDITABLE - no editors added)
                Word.Range cell41Range = table.Cell(4, 1).Range;
                cell41Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell41Range.Text = "Source: XYZ Research";
                cell41Range.Font.Bold = 0;
                cell41Range.Font.Size = 8;
               // cell41Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                Word.Range cell42Range = table.Cell(4, 2).Range;
                cell42Range.MoveEnd(Word.WdUnits.wdCharacter, -1);
                cell42Range.Text = "Source: XYZ Research";
                cell42Range.Font.Bold = 0;
                cell42Range.Font.Size = 8;
                //cell42Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Step 7: Add editable text after the table (inside Inner Control)
                //Word.Range afterTableRange = doc.Range(table.Range.End, table.Range.End);
                //afterTableRange.Text = "\r\nThis text is after the table, still inside the Inner Control.";
                //afterTableRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                // Move cursor after the master control
                Word.Range afterRange = masterControl.Range;
              //  afterRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
               // afterRange.Select();
                // Remove editors from last two rows (rows 3 and 4) only
                //for (int row = 3; row <= 4; row++)
                //{
                //    for (int col = 1; col <= 2; col++)
                //    {
                //        Word.Editors editors = table.Cell(row, col).Range.Editors;
                //        Word.Range rmCellrange = editors.Item(1).Range;
                //        rmCellrange.Editors.Item(1).Delete();
                //        for (int i = editors.Count; i >= 1; i--)
                //        {
                //            rmCellrange.Editors.Item(i).Delete();
                //        }
                //    }
                //}

                for (int row = 1; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = table.Cell(row, col).Range;
                        cellRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                    }
                }

                // STEP 8: Remove editors from rows 3 and 4
                for (int row = 3; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = table.Cell(row, col).Range;
                        while (cellRange.Editors.Count > 0)
                        {
                            cellRange.Editors.Item(1).Delete();
                        }
                    }
                }

                TraceLog("CreateNestedControls completed successfully");
            }
            //catch
            //{
            //    TraceLog("Error occurred while creating nested controls");

            //} 
            finally
            {
                // Protect the document with AllowOnlyReading to enforce editable regions
                doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, Password: "test");
            }
}

        public void ReproAddRemoveEditor()
        
        {
            Word.Document doc = this.Application.ActiveDocument;
            Word.Range range = this.Application.Selection.Range;

            if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
            {
                doc.Unprotect(Password: "test");
            }

            try
            {
                // Step 1: Create Master Control
                range.Text = " ";
                Word.ContentControl masterControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText, range);
                masterControl.Title = "Master Control";
                masterControl.Tag= "boxtype"; //charts##3294615616
                masterControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;

                // Step 2: Add text and enter
                Word.Range masterRange = masterControl.Range;
                // masterRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                //masterRange.Text=
                masterRange.Select();
                // Type =rand() and press Enter to execute it
                

                this.Application.Selection.TypeText("=rand()");  // 2 paragraphs, 3 sentences each
                this.Application.Activate();
                System.Windows.Forms.Application.DoEvents();

                System.Windows.Forms.SendKeys.Send("{ENTER}");
                System.Windows.Forms.Application.DoEvents();

                masterControl.Range.Select();
               // this.Application.Selection.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                //Select the second paragraph and insert a table

                Word.Paragraph secondParagraph = doc.Paragraphs[2];

                // Get the range of the second paragraph
                Word.Range paragraphRange = secondParagraph.Range;

                // Collapse to the end of the paragraph
                paragraphRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // Move back 1 character to be before the paragraph mark (optional)
                 paragraphRange.MoveEnd(Word.WdUnits.wdCharacter, -1);

                // Select the position
                paragraphRange.Select();

               // return;

                //masterRange.Text = "This is some random text inside the Master Control.\r\n";
                
                //masterRange.Text = "This is some random text inside the Master Control.\r\n";
                //Word.Range masterTextRange = doc.Range(masterControl.Range.Start, masterControl.Range.End - 1);
               // masterTextRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                //masterRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
               // masterRange.Select();
                

                // Step 3: Create Inner Control - Commented as it has to be added later
                //Word.Range innerRange = this.Application.Selection.Range;
                //innerRange.Text = " ";
                //Word.ContentControl innerControl = doc.ContentControls.Add(
                //    Word.WdContentControlType.wdContentControlRichText, innerRange);
                //innerControl.Title = "Inner Control";
                //innerControl.Tag = "Innter"; //charts##3294615616
                //innerControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;
                
                // Step 4: Add text and enter
               // Word.Range innerContentRange = innerControl.Range;
               // innerContentRange = doc.Range(masterControl.Range.Start, masterControl.Range.End - 1);

                //innerContentRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                //innerContentRange.Editors.Item(1).Range.Text = "This is some random text inside the Inner Control.\r\n";
                
                //innerContentRange.Text = "This is some random text inside the Inner Control.\r\n";
                //Word.Range innerEditableRange = doc.Range(innerControl.Range.Start, innerControl.Range.End - 1);
                //innerEditableRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                //innerContentRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                //innerContentRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                //innerContentRange.Select();
                
                // Step 5: Add table
                Word.Range tableRange = this.Application.Selection.Range;
                Word.Table table = doc.Tables.Add(tableRange, 4, 2);
                table.Borders.Enable = 1;

                // STEP 6: SET TEXT FIRST (before adding editors)
                // Row 1
                table.Cell(1, 1).Range.Text = "Figure 1. Test image 1\a";
                table.Cell(1, 2).Range.Text = "Figure 2. Test image 2\a";

                // Row 2
                table.Cell(2, 1).Range.Text = "[Image Placeholder 1]\a";
                table.Cell(2, 2).Range.Text = "[Image Placeholder 2]\a";

                // Row 3
                table.Cell(3, 1).Range.Text = "© 2025 XYZ Inc.\a";
                table.Cell(3, 2).Range.Text = "© 2025 XYZ Inc.\a";

                // Row 4
                table.Cell(4, 1).Range.Text = "Source: XYZ Research\a";
                table.Cell(4, 2).Range.Text = "Source: XYZ Research\a";

                // STEP 7: NOW add editors AFTER text is set
                for (int row = 1; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = table.Cell(row, col).Range;
                        cellRange.Editors.Add(Word.WdEditorType.wdEditorEveryone);
                    }
                }

                // STEP 8: Remove editors from rows 3 and 4
                for (int row = 3; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = table.Cell(row, col).Range;
                        while (cellRange.Editors.Count > 0)
                        {
                            cellRange.Editors.Item(1).Delete();
                        }
                    }
                }

                //Add a wrapper content control for table
                table.Range.Select();
                Word.Range innerRange = this.Application.Selection.Range;

                Word.ContentControl innerControl = doc.ContentControls.Add(
                    Word.WdContentControlType.wdContentControlRichText, innerRange);
                innerControl.Title = "Inner Control";
                innerControl.Tag = "InnerCC"; //charts##3294615616
                innerControl.Appearance = Word.WdContentControlAppearance.wdContentControlBoundingBox;

                innerControl.Range.Select();

                string wordOpenXML = this.Application.Selection.WordOpenXML;

                Word.Paragraph fourthParagraph = doc.Paragraphs[17];

                // Get the range of the fourth paragraph
                Word.Range fourthParagraphRange = fourthParagraph.Range;

                // Collapse to the end of the paragraph
                //fourthParagraphRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                // Move back 1 character to be before the paragraph mark (optional)
                fourthParagraphRange.MoveEnd(Word.WdUnits.wdCharacter, -1);

                fourthParagraphRange.Select();
                this.Application.Selection.Range.InsertXML(wordOpenXML);

                Word.Table firstTable = doc.Tables[1];

                for (int row = 3; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                          Word.Range cellRange = firstTable.Cell(row, col).Range;
                        while (cellRange.Editors.Count > 0)
                        {
                            cellRange.Editors.Item(1).Delete();
                        }
                    }
                }


                Word.Table secondTable = doc.Tables[2];

                for (int row = 3; row <= 4; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        Word.Range cellRange = secondTable.Cell(row, col).Range;
                        while (cellRange.Editors.Count > 0)
                        {
                            cellRange.Editors.Item(1).Delete();
                        }
                    }
                }


                //--Inserting random text inside the inner control to repro the issue

                //innerControl.Range.Select();
                //this.Application.Selection.TypeText("=rand()");  // 2 paragraphs, 3 sentences each
                //this.Application.Activate();
                //System.Windows.Forms.Application.DoEvents();

                //System.Windows.Forms.SendKeys.Send("{ENTER}");
                //System.Windows.Forms.Application.DoEvents();

                //-----Inserting random text inside the inner control to repro the issue

                //innerControl.Range.Select();
                //this.Application.Selection.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

              //  masterControl.Range.Select();
                //  this.Application.Selection.Range.Editors.Add(Word.WdEditorType.wdEditorEveryone);

                TraceLog("ReproAddRemoveEditor completed successfully");


            }
            finally
            {
                doc.Protect(Word.WdProtectionType.wdAllowOnlyReading, Password: "test");
            }
}
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
