using System.Collections.Concurrent;
using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Freezes panes on the worksheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        internal void FreezeCore(int topRows, int leftCols) {
            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteWorksheetView(worksheet => {
                OfficeXmlNode views = worksheet.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetViews");
                if (views.IsNone) {
                    if (topRows == 0 && leftCols == 0) return false;
                    views = InsertWorksheetViews(worksheet);
                }

                SetFrozenPanes(views, topRows, leftCols);
                return true;
            }, worksheetOnly: true);
        }

        private static void SetFrozenPanes(OfficeXmlNode views, int topRows, int leftCols) {
            if (topRows == 0 && leftCols == 0) {
                foreach (OfficeXmlNode view in views.Children(OfficeXmlNamespaces.SpreadsheetML, "sheetView")) {
                    view.RemoveChildren(OfficeXmlNamespaces.SpreadsheetML, "pane");
                    view.RemoveChildren(OfficeXmlNamespaces.SpreadsheetML, "selection");
                    AddSelection(view, null, "A1", "A1");
                }
                return;
            }

            OfficeXmlNode primaryView = GetOrCreateView(views);
            primaryView.RemoveChildren(OfficeXmlNamespaces.SpreadsheetML, "pane");
            primaryView.RemoveChildren(OfficeXmlNamespaces.SpreadsheetML, "selection");

            string topLeftCell = ExcelA1.CellReference(topRows + 1, leftCols + 1);
            string activePane = topRows > 0
                ? leftCols > 0 ? "bottomRight" : "bottomLeft"
                : "topRight";
            OfficeXmlNode pane = primaryView.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "pane");
            if (leftCols > 0) pane.SetInt32Attribute("xSplit", leftCols);
            if (topRows > 0) pane.SetInt32Attribute("ySplit", topRows);
            pane.SetAttribute("topLeftCell", topLeftCell);
            pane.SetAttribute("activePane", activePane);
            pane.SetAttribute("state", "frozen");
            primaryView.InsertBefore(pane, primaryView.FirstChild);

            if (topRows > 0 && leftCols > 0) {
                AddSelection(primaryView, "topRight", topLeftCell, topLeftCell);
                AddSelection(primaryView, null, "A1", "A1");
                AddSelection(primaryView, "bottomLeft", topLeftCell, topLeftCell);
                AddSelection(primaryView, "bottomRight", topLeftCell, topLeftCell);
            } else {
                AddSelection(primaryView, activePane, topLeftCell, topLeftCell);
                AddSelection(primaryView, null, "A1", "A1");
            }
        }

        private void WriteWorksheetView(Func<OfficeXmlNode, bool> edit, bool save = false, bool worksheetOnly = false) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                if (_excelDocument.TryEditDirectWorksheetView(this, markup => EditWorksheetViewsMarkup(markup, edit))) return;
                if (!worksheetOnly && _excelDocument.HasDeferredDirectDataSetImport) {
                    _excelDocument.MaterializeDeferredDataSetImport();
                }
                if (DeferredEntry is { } entry && _excelDocument.TryEditDeferredWorksheetView(entry, edit)) return;
                if (!edit(ViewWorksheetRoot)) return;
                if (save) ExcelPartRef.WorksheetForMetadata(this).SaveRoot();
                if (worksheetOnly) MarkRequiresWorksheetPreparation();
                else MarkRequiresSavePreparation();
            });
        }

        private static void AddSelection(OfficeXmlNode view, string? pane, string activeCell, string references) {
            OfficeXmlNode selection = view.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "selection");
            selection.SetAttribute("pane", pane);
            selection.SetAttribute("activeCell", activeCell);
            selection.SetAttribute("sqref", references);
            InsertViewChild(view, selection);
        }

        /// <summary>
        /// Shows or hides gridlines on the current sheet (view-level setting).
        /// </summary>
        internal void SetGridlinesVisibleCore(bool visible) {
            WriteWorksheetView(worksheet => {
                GetOrCreateSheetView(worksheet).SetBooleanAttribute("showGridLines", visible);
                return true;
            });
        }

        /// <summary>
        /// Gets whether worksheet row and column headings are displayed in the worksheet view.
        /// </summary>
        internal bool RowColumnHeadingsVisibleCore => ReadViewInfo().ShowRowColumnHeadings;

        /// <summary>
        /// Shows or hides worksheet row and column headings in the worksheet view.
        /// </summary>
        internal void SetRowColumnHeadingsVisibleCore(bool visible) {
            WriteWorksheetView(worksheet => {
                GetOrCreateSheetView(worksheet).SetBooleanAttribute("showRowColHeaders", visible);
                return true;
            });
        }

        /// <summary>
        /// Gets whether zero values are displayed in the worksheet view.
        /// </summary>
        internal bool ZeroValuesVisibleCore => ReadViewInfo().ShowZeros;

        /// <summary>
        /// Shows or hides zero values in the worksheet view.
        /// </summary>
        internal void SetZeroValuesVisibleCore(bool visible) {
            WriteWorksheetView(worksheet => {
                GetOrCreateSheetView(worksheet).SetBooleanAttribute("showZeros", visible);
                return true;
            });
        }

        /// <summary>
        /// Gets whether the worksheet view is displayed from right to left.
        /// </summary>
        internal bool RightToLeftCore => ReadViewInfo().RightToLeft;

        /// <summary>
        /// Sets whether the worksheet view is displayed from right to left.
        /// </summary>
        /// <param name="rightToLeft">True to display the worksheet from right to left; otherwise false.</param>
        internal void SetRightToLeftCore(bool rightToLeft) {
            WriteWorksheetView(worksheet => {
                GetOrCreateSheetView(worksheet).SetBooleanAttribute("rightToLeft", rightToLeft);
                return true;
            });
        }

        /// <summary>
        /// Gets the worksheet view zoom scale percentage, or null when no explicit zoom is configured.
        /// </summary>
        internal uint? GetZoomScaleCore() => ReadViewInfo().ZoomScale;

        /// <summary>
        /// Sets the worksheet view zoom scale percentage.
        /// </summary>
        /// <param name="scale">Zoom percentage from 10 through 400.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        internal void SetZoomScaleCore(uint scale, bool save = true) {
            if (scale < 10U || scale > 400U) {
                throw new ArgumentOutOfRangeException(nameof(scale), "Worksheet zoom scale must be between 10 and 400 percent.");
            }

            WriteWorksheetView(worksheet => {
                GetOrCreateSheetView(worksheet).SetUInt32Attribute("zoomScale", scale);
                return true;
            }, save);
        }

        internal void SetWorksheetSelection(string activeCell, IReadOnlyList<string> selectedRanges, string? pane = null, bool save = true) {
            if (string.IsNullOrWhiteSpace(activeCell)) throw new ArgumentNullException(nameof(activeCell));
            ArgumentNullException.ThrowIfNull(selectedRanges);

            string sequenceOfReferences = selectedRanges.Count == 0
                ? activeCell
                : string.Join(" ", selectedRanges.Where(range => !string.IsNullOrWhiteSpace(range)));
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                sequenceOfReferences = activeCell;
            }

            WriteWorksheetView(worksheet => {
                OfficeXmlNode view = GetOrCreateSheetView(worksheet);
                foreach (OfficeXmlNode existing in view.Children(OfficeXmlNamespaces.SpreadsheetML, "selection").ToList()) {
                    if (existing.Attribute("pane") == pane) existing.Remove();
                }

                OfficeXmlNode selection = view.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "selection");
                selection.SetAttribute("activeCell", activeCell);
                selection.SetAttribute("sqref", sequenceOfReferences);
                selection.SetAttribute("pane", pane);
                InsertViewChild(view, selection);
                return true;
            }, save);
        }

        private static OfficeXmlNode GetOrCreateSheetView(OfficeXmlNode worksheet) {
            OfficeXmlNode views = worksheet.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetViews");
            return GetOrCreateView(views.IsNone ? InsertWorksheetViews(worksheet) : views);
        }

        private static OfficeXmlNode InsertWorksheetViews(OfficeXmlNode worksheet) {
            OfficeXmlNode next = worksheet.FirstChild;
            while (next.Is(OfficeXmlNamespaces.SpreadsheetML, "sheetPr")
                || next.Is(OfficeXmlNamespaces.SpreadsheetML, "dimension")) {
                next = next.NextSibling;
            }

            return worksheet.InsertBefore(worksheet.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "sheetViews"), next);
        }

        private static OfficeXmlNode GetOrCreateView(OfficeXmlNode views) {
            OfficeXmlNode view = views.Child(OfficeXmlNamespaces.SpreadsheetML, "sheetView");
            if (view.IsNone) {
                view = InsertViewChild(views, views.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "sheetView"));
            }
            if (!view.HasAttribute("workbookViewId")) view.SetUInt32Attribute("workbookViewId", 0U);
            return view;
        }

        private static OfficeXmlNode InsertViewChild(OfficeXmlNode parent, OfficeXmlNode child) {
            OfficeXmlNode next = parent.Child(OfficeXmlNamespaces.SpreadsheetML, "extLst");
            if (child.Is(OfficeXmlNamespaces.SpreadsheetML, "selection")) {
                foreach (OfficeXmlNode sibling in parent.Children()) {
                    if (sibling.Is(OfficeXmlNamespaces.SpreadsheetML, "pivotSelection") || sibling == next) {
                        next = sibling;
                        break;
                    }
                }
            }
            return parent.InsertBefore(child, next);
        }

    }
}
