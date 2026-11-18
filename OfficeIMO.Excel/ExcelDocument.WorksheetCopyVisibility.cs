using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Gives a copied worksheet the hidden or very hidden state of its source. The state lives on the workbook's
        /// <c>sheet</c> entry, not in the worksheet part, so a copy of the part alone is always visible.
        /// </summary>
        private static void CopyWorksheetVisibility(Sheet sourceSheet, Sheet copiedSheet) {
            copiedSheet.State = IsHiddenSheet(sourceSheet) ? sourceSheet.State!.Value : (SheetStateValues?)null;
        }

        /// <summary>
        /// Keeps a workbook openable after copies that kept their source's hidden state. A workbook needs a visible
        /// sheet and a visible active sheet, so when no sheet is visible the first copied worksheet is shown, and when
        /// the active sheet is hidden the first visible worksheet becomes active.
        /// </summary>
        private void EnsureVisibleActiveWorksheetAfterCopy(ExcelSheet firstCopiedSheet) {
            ExcelSheet? activate = Locking.ExecuteWrite<ExcelSheet?>(EnsureLock(), () => {
                List<Sheet> tabs = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                if (tabs.All(IsHiddenSheet)) {
                    firstCopiedSheet.SheetElement.State = null;
                    WorkbookRoot.Save();
                }

                int? activeTab = GetActiveWorksheetIndex(WorkbookRoot, tabs.Count);
                if (!activeTab.HasValue || !IsHiddenSheet(tabs[activeTab.Value])) {
                    return null;
                }

                ExcelSheet? visibleWorksheet = GetSheetsForLockedOperation().FirstOrDefault(_ => !_.Hidden);
                if (visibleWorksheet == null) {
                    // Only a non-worksheet tab, such as a chart sheet, is visible: open on it and leave the copies hidden.
                    int visibleTab = tabs.FindIndex(_ => !IsHiddenSheet(_));
                    WorkbookView workbookView = GetOrCreatePrimaryWorkbookViewForActiveWorksheet();
                    workbookView.ActiveTab = (uint)visibleTab;
                    workbookView.FirstSheet = (uint)visibleTab;
                    WorkbookRoot.Save();
                    MarkPackageDirty();
                }

                return visibleWorksheet;
            });

            if (activate != null) {
                SetActiveWorksheet(activate);
            }
        }
    
        /// <summary>
        /// Moves the activation off a worksheet that was just hidden: to the nearest visible worksheet on its right,
        /// otherwise on its left. Leaves the workbook unchanged when the hidden worksheet was not active or no
        /// worksheet is visible.
        /// </summary>
        /// <summary>
        /// Activates a visible worksheet when the active one is hidden. Used after imports whose source workbook
        /// view is not carried over, so the implicit first tab may be a hidden sheet.
        /// </summary>
        internal void EnsureActiveWorksheetVisible() {
            ExcelSheet? active = ActiveWorksheet;
            if (active != null && active.Hidden) {
                ActivateVisibleWorksheetAfterHiding(active);
            }
        }

        internal void ActivateVisibleWorksheetAfterHiding(ExcelSheet hiddenSheet) {
            ExcelSheet? activate = Locking.ExecuteRead(EnsureLock(), () => {
                List<Sheet> tabs = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                int hiddenTab = tabs.FindIndex(tab => ReferenceEquals(tab, hiddenSheet.SheetElement));
                if (hiddenTab < 0 || GetActiveWorksheetIndex(WorkbookRoot, tabs.Count) != hiddenTab) {
                    return null;
                }

                List<ExcelSheet> worksheets = GetSheetsForLockedOperation();
                int TabOf(ExcelSheet sheet) => tabs.FindIndex(tab => ReferenceEquals(tab, sheet.SheetElement));
                return worksheets.Where(sheet => !sheet.Hidden && TabOf(sheet) > hiddenTab).OrderBy(TabOf).FirstOrDefault()
                    ?? worksheets.Where(sheet => !sheet.Hidden && TabOf(sheet) < hiddenTab).OrderByDescending(TabOf).FirstOrDefault();
            });

            if (activate != null) {
                SetActiveWorksheet(activate);
            }
        }
    }
}
