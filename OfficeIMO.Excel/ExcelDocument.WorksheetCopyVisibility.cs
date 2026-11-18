using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Column definitions the library created with the worksheet default width. Only these follow a later
        /// default-width change or count as empty when cleared; a loaded width that merely equals the default stays as
        /// authored. The table holds no strong references, so removed definitions are released with the DOM.
        /// </summary>
        internal System.Runtime.CompilerServices.ConditionalWeakTable<Column, object> DefaultWidthColumns { get; } = new();

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
        /// Moves every workbook view whose active tab is hidden to the nearest visible tab on its right, otherwise on
        /// its left, as Excel does when the active sheet is hidden. A workbook without views has the first tab active
        /// implicitly, so a primary view is created when that tab is hidden. Views keep their active tab when no tab
        /// is visible. Callers hold the document write lock, so the hide and the new activation are one change.
        /// </summary>
        /// <returns>True when a view changed.</returns>
        internal bool MoveActiveTabsOffHiddenSheetsCore() {
            List<Sheet> tabs = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            if (tabs.Count == 0) return false;

            List<WorkbookView> views = WorkbookRoot.GetFirstChild<BookViews>()?.Elements<WorkbookView>().ToList() ?? new List<WorkbookView>();
            if (views.Count == 0) {
                if (!IsHiddenSheet(tabs[0]) || FindVisibleTab(tabs, 0) < 0) return false;
                views.Add(GetOrCreatePrimaryWorkbookViewForActiveWorksheet());
            }

            bool changed = false;
            int primaryActive = -1;
            for (int viewIndex = 0; viewIndex < views.Count; viewIndex++) {
                WorkbookView view = views[viewIndex];
                int active = (int)Math.Min(view.ActiveTab?.Value ?? 0U, (uint)(tabs.Count - 1));
                if (!IsHiddenSheet(tabs[active])) continue;
                int replacement = FindVisibleTab(tabs, active);
                if (replacement < 0) continue;
                view.ActiveTab = (uint)replacement;
                view.FirstSheet = (uint)replacement;
                if (viewIndex == 0) primaryActive = replacement;
                changed = true;
            }

            if (primaryActive >= 0) {
                for (int index = 0; index < tabs.Count; index++) {
                    string? relationshipId = tabs[index].Id?.Value;
                    if (relationshipId == null || WorkbookPartRoot.GetPartById(relationshipId) is not WorksheetPart worksheetPart) continue;
                    Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                    SheetView sheetView = GetOrCreatePrimarySheetViewForActiveWorksheet(worksheet);
                    sheetView.WorkbookViewId ??= 0U;
                    sheetView.TabSelected = index == primaryActive;
                    worksheet.Save();
                }
            }

            if (changed) {
                WorkbookRoot.Save();
                MarkPackageDirty();
            }

            return changed;
        }

        /// <summary>Used after imports whose source workbook view is not carried over.</summary>
        internal void EnsureActiveTabVisible() =>
            Locking.ExecuteWrite(EnsureLock(), () => { MoveActiveTabsOffHiddenSheetsCore(); });

        private static int FindVisibleTab(IReadOnlyList<Sheet> tabs, int hiddenTab) {
            for (int index = hiddenTab + 1; index < tabs.Count; index++) {
                if (!IsHiddenSheet(tabs[index])) return index;
            }

            for (int index = hiddenTab - 1; index >= 0; index--) {
                if (!IsHiddenSheet(tabs[index])) return index;
            }

            return -1;
        }
}
}
