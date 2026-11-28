using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupSheetViewArtifacts() => CleanupSheetViewArtifacts(ViewWorksheetRoot);

        internal static void CleanupSheetViewArtifacts(OfficeXmlNode worksheet) {
            OfficeXmlNode views = worksheet.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetViews");
            if (views.IsNone) return;

            var sheetViews = views.Children(OfficeXmlNamespaces.SpreadsheetML, "sheetView").ToList();
            foreach (OfficeXmlNode view in sheetViews) CleanupSheetView(view);
            if (sheetViews.Count == 0) views.Remove();
        }

        private static void CleanupSheetView(OfficeXmlNode view) {
            if (!view.HasAttribute("workbookViewId")) view.SetUInt32Attribute("workbookViewId", 0U);

            OfficeXmlNode pane = view.Child(OfficeXmlNamespaces.SpreadsheetML, "pane");
            string[] frozenPanes = Array.Empty<string>();
            string paneCell = "A1";
            string? state = pane.Attribute("state");
            if (state is "frozen" or "frozenSplit") {
                int frozenRows = NormalizeFrozenSplit(pane.DoubleAttribute("ySplit"));
                int frozenColumns = NormalizeFrozenSplit(pane.DoubleAttribute("xSplit"));
                pane.SetInt32Attribute("ySplit", frozenRows > 0 ? frozenRows : null);
                pane.SetInt32Attribute("xSplit", frozenColumns > 0 ? frozenColumns : null);

                if (frozenRows <= 0 && frozenColumns <= 0) {
                    pane.Remove();
                } else {
                    paneCell = ExcelA1.CellReference(frozenRows + 1, frozenColumns + 1);
                    pane.SetAttribute("topLeftCell", paneCell);
                    frozenPanes = frozenRows > 0
                        ? frozenColumns > 0 ? new[] { "topRight", "bottomLeft", "bottomRight" } : new[] { "bottomLeft" }
                        : new[] { "topRight" };
                    pane.SetAttribute("activePane", frozenPanes[frozenPanes.Length - 1]);
                }
            } else if (pane.Exists && !IsValidCellReference(pane.Attribute("topLeftCell"))) {
                pane.SetAttribute("topLeftCell", "A1");
            }

            var seenPanes = new HashSet<string>(StringComparer.Ordinal);
            bool hasUnqualifiedSelection = false;
            foreach (OfficeXmlNode selection in view.Children(OfficeXmlNamespaces.SpreadsheetML, "selection").ToList()) {
                string? selectionPane = selection.Attribute("pane");
                if (selectionPane == null) {
                    if (hasUnqualifiedSelection) {
                        selection.Remove();
                    } else {
                        hasUnqualifiedSelection = true;
                        NormalizeSelection(selection, "A1");
                    }
                } else if (!frozenPanes.Contains(selectionPane) || !seenPanes.Add(selectionPane)) {
                    selection.Remove();
                } else {
                    NormalizeSelection(selection, paneCell);
                }
            }

            foreach (string frozenPane in frozenPanes) {
                if (!seenPanes.Contains(frozenPane)) AddSelection(view, frozenPane, paneCell, paneCell);
            }
            if (!hasUnqualifiedSelection) AddSelection(view, null, "A1", "A1");
        }

        private static int NormalizeFrozenSplit(double? split) =>
            split is > 0D ? (int)Math.Truncate(split.Value) : 0;

        private static bool IsValidCellReference(string? reference) {
            var (row, column) = ExcelA1.ParseCellRef(reference ?? string.Empty);
            return row > 0 && column > 0;
        }

        private static void NormalizeSelection(OfficeXmlNode selection, string defaultCell) {
            if (!IsValidCellReference(selection.Attribute("activeCell"))) {
                selection.SetAttribute("activeCell", defaultCell);
            }

            string sqref = selection.Attribute("sqref") ?? string.Empty;
            if (!IsValidSelectionReferenceList(sqref)) {
                selection.SetAttribute("sqref", selection.Attribute("activeCell") ?? defaultCell);
            }
        }

        private static bool IsValidSelectionReferenceList(string? referenceList) {
            if (string.IsNullOrWhiteSpace(referenceList)) {
                return false;
            }

            foreach (string reference in referenceList!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (IsValidCellReference(reference)) {
                    continue;
                }

                if (ExcelA1.TryParseRange(reference, out _, out _, out _, out _)) {
                    continue;
                }

                return false;
            }

            return true;
        }

    }
}
