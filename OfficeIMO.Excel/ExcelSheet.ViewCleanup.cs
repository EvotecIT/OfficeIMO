using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupSheetViewArtifacts() {
            var worksheet = WorksheetRoot;
            var sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                return;
            }

            foreach (var sheetView in sheetViews.Elements<SheetView>().ToList()) {
                CleanupSheetView(sheetView);
            }

            if (!sheetViews.Elements<SheetView>().Any()) {
                worksheet.RemoveChild(sheetViews);
            }
        }

        private void CleanupSheetView(SheetView sheetView) {
            sheetView.WorkbookViewId ??= 0U;

            var pane = sheetView.GetFirstChild<Pane>();
            List<PaneValues> expectedFrozenPanes = new();
            string paneCell = "A1";

            if (pane != null && IsFrozenPaneState(pane.State?.Value)) {
                int frozenRows = NormalizeFrozenSplit(pane.VerticalSplit);
                int frozenColumns = NormalizeFrozenSplit(pane.HorizontalSplit);

                if (frozenRows <= 0) {
                    pane.VerticalSplit = null;
                } else {
                    pane.VerticalSplit = frozenRows;
                }

                if (frozenColumns <= 0) {
                    pane.HorizontalSplit = null;
                } else {
                    pane.HorizontalSplit = frozenColumns;
                }

                if (frozenRows <= 0 && frozenColumns <= 0) {
                    pane.Remove();
                    pane = null;
                } else {
                    paneCell = A1.ColumnIndexToLetters(frozenColumns + 1) + (frozenRows + 1).ToString(CultureInfo.InvariantCulture);
                    pane.TopLeftCell = paneCell;

                    if (frozenRows > 0 && frozenColumns > 0) {
                        expectedFrozenPanes.Add(PaneValues.TopRight);
                        expectedFrozenPanes.Add(PaneValues.BottomLeft);
                        expectedFrozenPanes.Add(PaneValues.BottomRight);
                        pane.ActivePane = PaneValues.BottomRight;
                    } else if (frozenRows > 0) {
                        expectedFrozenPanes.Add(PaneValues.BottomLeft);
                        pane.ActivePane = PaneValues.BottomLeft;
                    } else {
                        expectedFrozenPanes.Add(PaneValues.TopRight);
                        pane.ActivePane = PaneValues.TopRight;
                    }
                }
            } else if (pane != null) {
                if (!IsValidCellReference(pane.TopLeftCell?.Value)) {
                    pane.TopLeftCell = "A1";
                }
                paneCell = pane.TopLeftCell?.Value ?? "A1";
            }

            if (pane == null) {
                foreach (var selection in sheetView.Elements<Selection>().Where(selection => selection.Pane != null).ToList()) {
                    selection.Remove();
                }
            }

            var seenKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (var selection in sheetView.Elements<Selection>().ToList()) {
                string key = selection.Pane?.Value.ToString() ?? "<default>";
                if (!seenKeys.Add(key)) {
                    selection.Remove();
                    continue;
                }

                if (selection.Pane != null) {
                    if (pane == null || !expectedFrozenPanes.Contains(selection.Pane.Value)) {
                        selection.Remove();
                        continue;
                    }

                    string frozenCell = pane?.TopLeftCell?.Value ?? paneCell;
                    NormalizeSelection(selection, frozenCell);
                } else {
                    NormalizeSelection(selection, "A1");
                }
            }

            if (pane != null && expectedFrozenPanes.Count > 0) {
                foreach (var paneValue in expectedFrozenPanes) {
                    if (sheetView.Elements<Selection>().Any(selection => selection.Pane?.Value == paneValue)) {
                        continue;
                    }

                    sheetView.Append(CreateSelection(paneValue, pane.TopLeftCell?.Value ?? paneCell));
                }
            }

            if (!sheetView.Elements<Selection>().Any(selection => selection.Pane == null)) {
                sheetView.Append(CreateSelection(null, "A1"));
            }
        }

        private static bool IsFrozenPaneState(PaneStateValues? state) {
            return state == PaneStateValues.Frozen || state == PaneStateValues.FrozenSplit;
        }

        private static int NormalizeFrozenSplit(DoubleValue? split) {
            if (split?.Value == null || split.Value <= 0) {
                return 0;
            }

            return (int)Math.Truncate(split.Value);
        }

        private static bool IsValidCellReference(string? reference) {
            var (row, column) = A1.ParseCellRef(reference ?? string.Empty);
            return row > 0 && column > 0;
        }

        private static void NormalizeSelection(Selection selection, string defaultCell) {
            if (!IsValidCellReference(selection.ActiveCell?.Value)) {
                selection.ActiveCell = defaultCell;
            }

            string sqref = selection.SequenceOfReferences?.InnerText ?? string.Empty;
            if (!IsValidCellReference(sqref)) {
                selection.SequenceOfReferences = new ListValue<StringValue> { InnerText = selection.ActiveCell?.Value ?? defaultCell };
            }
        }

        private static Selection CreateSelection(PaneValues? paneValue, string cellReference) {
            var selection = new Selection {
                ActiveCell = cellReference,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = cellReference }
            };

            if (paneValue != null) {
                selection.Pane = paneValue.Value;
            }

            return selection;
        }
    }
}
