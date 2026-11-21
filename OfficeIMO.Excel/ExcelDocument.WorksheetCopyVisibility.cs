using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Column definitions the library created with the worksheet default width. Only these follow a later
        /// default-width change or count as empty when cleared; a loaded width that merely equals the default stays as
        /// authored. The table holds no strong references, so removed definitions are released with the DOM.
        /// </summary>
        internal System.Runtime.CompilerServices.ConditionalWeakTable<Column, AssignedColumnWidth> DefaultWidthColumns { get; } = new();

        /// <summary>The width the library last assigned to a column it created.</summary>
        internal sealed class AssignedColumnWidth {
            internal AssignedColumnWidth(double width) => Width = width;
            internal double Width { get; set; }
        }

        /// <summary>
        /// Records the library-created default-width columns of every loaded worksheet by part and position, so a
        /// rollback that rebuilds the worksheet DOM can register the restored columns again.
        /// </summary>
        internal Dictionary<Uri, List<(int Position, uint? Min, uint? Max, double Width)>> CaptureDefaultWidthColumns() {
            var state = new Dictionary<Uri, List<(int Position, uint? Min, uint? Max, double Width)>>();
            foreach (WorksheetPart part in WorkbookPartRoot.WorksheetParts) {
                if (!part.IsRootElementLoaded) continue;
                List<Column>? columns = part.Worksheet?.GetFirstChild<Columns>()?.Elements<Column>().ToList();
                if (columns == null) continue;
                for (int position = 0; position < columns.Count; position++) {
                    if (!DefaultWidthColumns.TryGetValue(columns[position], out AssignedColumnWidth? assigned)) continue;
                    if (!state.TryGetValue(part.Uri, out var entries)) state[part.Uri] = entries = new();
                    entries.Add((position, columns[position].Min?.Value, columns[position].Max?.Value, assigned.Width));
                }
            }

            return state;
        }

        /// <summary>Registers the restored columns recorded by <see cref="CaptureDefaultWidthColumns"/>.</summary>
        internal void RestoreDefaultWidthColumns(Dictionary<Uri, List<(int Position, uint? Min, uint? Max, double Width)>> state) {
            if (state.Count == 0) return;
            foreach (WorksheetPart part in WorkbookPartRoot.WorksheetParts) {
                if (!state.TryGetValue(part.Uri, out var entries)) continue;
                List<Column>? columns = part.Worksheet?.GetFirstChild<Columns>()?.Elements<Column>().ToList();
                if (columns == null) continue;
                foreach (var entry in entries) {
                    if (entry.Position >= columns.Count) continue;
                    Column column = columns[entry.Position];
                    if (column.Min?.Value != entry.Min || column.Max?.Value != entry.Max || DefaultWidthColumns.TryGetValue(column, out _)) continue;
                    DefaultWidthColumns.Add(column, new AssignedColumnWidth(entry.Width));
                }
            }
        }

        /// <summary>
        /// Registers the copies of library-created default-width columns after a worksheet is cloned, so the copy keeps
        /// following the default width. Columns are paired by position and range; a mismatch registers nothing.
        /// </summary>
        internal void CarryDefaultWidthColumns(ExcelDocument source, Worksheet? sourceWorksheet, Worksheet? copiedWorksheet) {
            List<Column>? sourceColumns = sourceWorksheet?.GetFirstChild<Columns>()?.Elements<Column>().ToList();
            List<Column>? copiedColumns = copiedWorksheet?.GetFirstChild<Columns>()?.Elements<Column>().ToList();
            if (sourceColumns == null || copiedColumns == null || sourceColumns.Count != copiedColumns.Count) return;
            for (int index = 0; index < sourceColumns.Count; index++) {
                Column original = sourceColumns[index], copy = copiedColumns[index];
                if (source.DefaultWidthColumns.TryGetValue(original, out AssignedColumnWidth? assigned)
                    && original.Min?.Value == copy.Min?.Value && original.Max?.Value == copy.Max?.Value
                    && !DefaultWidthColumns.TryGetValue(copy, out _)) {
                    DefaultWidthColumns.Add(copy, new AssignedColumnWidth(assigned.Width));
                }
            }
        }

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
        /// Moves workbook views whose active tab is a hidden sheet to the nearest visible tab on its right, otherwise
        /// on its left, as Excel does when the active sheet is hidden. With <paramref name="hiddenTab"/> only views on
        /// that tab move (the public hide path); without it every view on a hidden tab moves (import repair). A
        /// workbook without views has the first tab active implicitly, so a primary view is created when needed. Views
        /// keep their tab when no other tab is visible. Each moved view's tab selection moves with it. Callers hold
        /// the document write lock, so the hide and the new activation are one change.
        /// </summary>
        /// <returns>True when a view changed.</returns>
        internal bool MoveActiveTabsOffHiddenSheetsCore(int? hiddenTab = null) {
            List<Sheet> tabs = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            if (tabs.Count == 0) return false;

            bool ShouldMove(int active) => IsHiddenSheet(tabs[active]) && (hiddenTab == null || hiddenTab == active);

            List<WorkbookView> views = WorkbookRoot.GetFirstChild<BookViews>()?.Elements<WorkbookView>().ToList() ?? new List<WorkbookView>();
            if (views.Count == 0) {
                if (!ShouldMove(0) || FindVisibleTab(tabs, 0) < 0) return false;
                views.Add(GetOrCreatePrimaryWorkbookViewForActiveWorksheet());
            }

            bool changed = false;
            for (int viewIndex = 0; viewIndex < views.Count; viewIndex++) {
                WorkbookView view = views[viewIndex];
                int active = (int)Math.Min(view.ActiveTab?.Value ?? 0U, (uint)(tabs.Count - 1));
                if (!ShouldMove(active)) continue;
                int replacement = FindVisibleTab(tabs, active);
                if (replacement < 0) continue;
                view.ActiveTab = (uint)replacement;
                if ((view.FirstSheet?.Value ?? 0U) > (uint)replacement) view.FirstSheet = (uint)replacement;
                SetTabSelected(tabs[active], (uint)viewIndex, selected: false);
                SetTabSelected(tabs[replacement], (uint)viewIndex, selected: true);
                changed = true;
            }

            if (changed) {
                WorkbookRoot.Save();
                MarkPackageDirty();
            }

            return changed;
        }

        /// <summary>
        /// Sets the tab selection of one workbook view on a worksheet or chart sheet, matching its view by
        /// <c>workbookViewId</c>. A missing primary worksheet view is created in schema order only to select the tab.
        /// </summary>
        private void SetTabSelected(Sheet tab, uint workbookViewId, bool selected) {
            string? relationshipId = tab.Id?.Value;
            if (relationshipId == null) return;
            OpenXmlPart part = WorkbookPartRoot.GetPartById(relationshipId);
            if (part is WorksheetPart worksheetPart && worksheetPart.Worksheet is Worksheet worksheet) {
                SheetView? view = worksheet.GetFirstChild<SheetViews>()?.Elements<SheetView>()
                    .FirstOrDefault(candidate => (candidate.WorkbookViewId?.Value ?? 0U) == workbookViewId);
                if (view == null) {
                    // Only the primary window gets a new view; a per-window view for another window is left to the
                    // application, since several sheet views per worksheet limit the formats the workbook can be saved in.
                    if (!selected || workbookViewId != 0U) return;
                    SheetViews? views = worksheet.GetFirstChild<SheetViews>();
                    if (views == null) {
                        views = new SheetViews();
                        worksheet.AddChild(views, true);
                    }
                    view = new SheetView { WorkbookViewId = workbookViewId };
                    // Sheet views precede an optional extension list.
                    ExtensionList? extensions = views.GetFirstChild<ExtensionList>();
                    if (extensions != null) views.InsertBefore(view, extensions);
                    else views.Append(view);
                }

                if (selected) view.TabSelected = true;
                else if (view.TabSelected?.Value == true) view.TabSelected = null;
                worksheet.Save();
            } else if (part is ChartsheetPart chartsheetPart && chartsheetPart.Chartsheet is Chartsheet chartsheet) {
                ChartSheetView? view = chartsheet.GetFirstChild<ChartSheetViews>()?.Elements<ChartSheetView>()
                    .FirstOrDefault(candidate => (candidate.WorkbookViewId?.Value ?? 0U) == workbookViewId);
                if (view == null) return;
                if (selected) view.TabSelected = true;
                else if (view.TabSelected?.Value == true) view.TabSelected = null;
                chartsheet.Save();
            }
        }

        /// <summary>
        /// Applies an imported workbook view's active and first tab. Values at the implicit first tab are not
        /// written, so a workbook whose view is the default keeps no workbook view.
        /// </summary>
        internal void ApplyImportedActiveTab(uint? activeTab, uint? firstSheet, int sourceTabCount) {
            Locking.ExecuteWrite(EnsureLock(), () => {
                int count = WorkbookRoot.Sheets?.Elements<Sheet>().Count() ?? 0;
                // Source indexes only name the same tabs when every source tab was projected, in order.
                if (count == 0 || count != sourceTabCount) return;
                uint active = Math.Min(activeTab ?? 0U, (uint)(count - 1));
                uint first = Math.Min(firstSheet ?? 0U, (uint)(count - 1));
                if (active == 0U && first == 0U) return;
                WorkbookView view = GetOrCreatePrimaryWorkbookViewForActiveWorksheet();
                view.ActiveTab = active;
                if (first != 0U) view.FirstSheet = first;
                WorkbookRoot.Save();
            });
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
