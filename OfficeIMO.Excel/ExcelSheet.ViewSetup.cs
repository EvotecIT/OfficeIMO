using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
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

        internal void SetWorksheetSelection(string activeCell, IReadOnlyList<string> selectedRanges, PaneValues? pane = null, bool save = true) {
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
                string? paneValue = pane?.ToString();
                foreach (OfficeXmlNode existing in view.Children(OfficeXmlNamespaces.SpreadsheetML, "selection").ToList()) {
                    if (existing.Attribute("pane") == paneValue) existing.Remove();
                }

                OfficeXmlNode selection = view.CreateElement(OfficeXmlNamespaces.SpreadsheetML, "selection");
                selection.SetAttribute("activeCell", activeCell);
                selection.SetAttribute("sqref", sequenceOfReferences);
                selection.SetAttribute("pane", paneValue);
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

        /// <summary>
        /// Configures basic print/page setup for the sheet.
        /// </summary>
        /// <param name="fitToWidth">Number of pages to fit horizontally (1 = fit to one page).</param>
        /// <param name="fitToHeight">Number of pages to fit vertically (0 = unlimited).</param>
        /// <param name="scale">Manual scale (10-400). Ignored if FitToWidth/Height are specified.</param>
        /// <param name="pageOrder">Optional multi-page print order.</param>
        /// <param name="paperSize">Optional known paper size.</param>
        internal void SetPageSetupCore(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null, ExcelPageOrder? pageOrder = null, ExcelPaperSize? paperSize = null) {
            if (scale is < 10U or > 400U) {
                throw new ArgumentOutOfRangeException(nameof(scale), "Manual print scale must be between 10 and 400 percent.");
            }

            if (paperSize.HasValue) {
                ValidatePaperSize(paperSize.Value);
            }

            Persistence.OwnsTablePrintDefaults = false;
            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = GetOrCreatePageSetup(ws);

                if (fitToWidth != null) pageSetup.FitToWidth = fitToWidth.Value;
                if (fitToHeight != null) pageSetup.FitToHeight = fitToHeight.Value;
                if (scale != null) pageSetup.Scale = scale.Value;
                if (paperSize != null) pageSetup.PaperSize = (uint)paperSize.Value;
                if (pageOrder != null) {
                    pageSetup.PageOrder = pageOrder == ExcelPageOrder.OverThenDown
                        ? DocumentFormat.OpenXml.Spreadsheet.PageOrderValues.OverThenDown
                        : DocumentFormat.OpenXml.Spreadsheet.PageOrderValues.DownThenOver;
                }

                if (fitToWidth != null || fitToHeight != null) {
                    var sheetProperties = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>();
                    if (sheetProperties == null) {
                        sheetProperties = new DocumentFormat.OpenXml.Spreadsheet.SheetProperties();
                        ws.InsertAt(sheetProperties, 0);
                    }

                    var pageSetupProperties = sheetProperties.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties>();
                    if (pageSetupProperties == null) {
                        pageSetupProperties = new DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties();
                        sheetProperties.Append(pageSetupProperties);
                    }

                    pageSetupProperties.FitToPage = true;
                }

            });
        }

        internal void SetPageSetupAndClearStaleFit(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null, ExcelPageOrder? pageOrder = null, ExcelPaperSize? paperSize = null) {
            if (scale is < 10U or > 400U) {
                throw new ArgumentOutOfRangeException(nameof(scale), "Manual print scale must be between 10 and 400 percent.");
            }

            if (paperSize.HasValue) {
                ValidatePaperSize(paperSize.Value);
            }

            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = GetOrCreatePageSetup(ws);

                if (fitToWidth != null) pageSetup.FitToWidth = fitToWidth.Value;
                else {
                    pageSetup.FitToWidth = null;
                    pageSetup.RemoveAttribute("fitToWidth", string.Empty);
                }

                if (fitToHeight != null) pageSetup.FitToHeight = fitToHeight.Value;
                else {
                    pageSetup.FitToHeight = null;
                    pageSetup.RemoveAttribute("fitToHeight", string.Empty);
                }

                if (scale != null) pageSetup.Scale = scale.Value;
                else {
                    pageSetup.Scale = null;
                    pageSetup.RemoveAttribute("scale", string.Empty);
                }

                if (paperSize != null) pageSetup.PaperSize = (uint)paperSize.Value;

                if (pageOrder != null) {
                    pageSetup.PageOrder = pageOrder == ExcelPageOrder.OverThenDown
                        ? DocumentFormat.OpenXml.Spreadsheet.PageOrderValues.OverThenDown
                        : DocumentFormat.OpenXml.Spreadsheet.PageOrderValues.DownThenOver;
                }

                var sheetProperties = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetProperties>();
                var pageSetupProperties = sheetProperties?.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties>();
                if (fitToWidth != null || fitToHeight != null) {
                    if (sheetProperties == null) {
                        sheetProperties = new DocumentFormat.OpenXml.Spreadsheet.SheetProperties();
                        ws.InsertAt(sheetProperties, 0);
                    }

                    if (pageSetupProperties == null) {
                        pageSetupProperties = new DocumentFormat.OpenXml.Spreadsheet.PageSetupProperties();
                        sheetProperties.Append(pageSetupProperties);
                    }

                    pageSetupProperties.FitToPage = true;
                } else if (pageSetupProperties != null) {
                    pageSetupProperties.FitToPage = null;
                    pageSetupProperties.RemoveAttribute("fitToPage", string.Empty);
                }

            });
        }
    }
}
