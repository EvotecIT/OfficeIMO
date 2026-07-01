using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Freezes panes on the worksheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        public void Freeze(int topRows = 0, int leftCols = 0) {
            if (_excelDocument.TryApplyDirectWorksheetFreezeMetadata(this, topRows, leftCols)) {
                return;
            }

            using var preserveDirectDataSet = _excelDocument.PreserveDirectDataSetSaveCandidateDuringDirtyMarks();
            WriteLockWorksheetPreparationOnly(() => {
                Worksheet worksheet = WorksheetRoot;
                SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();

                if (topRows == 0 && leftCols == 0) {
                    if (sheetViews != null) {
                        worksheet.RemoveChild(sheetViews);
                        return true;
                    }
                    return false;
                }

                if (sheetViews == null) {
                    sheetViews = new SheetViews();

                    // Remove SheetData temporarily if it exists
                    var sheetData = worksheet.GetFirstChild<SheetData>();
                    if (sheetData != null) {
                        worksheet.RemoveChild(sheetData);
                    } else {
                        sheetData = new SheetData();
                    }

                    // Add sheetViews first
                    worksheet.AppendChild(sheetViews);

                    // Then add SheetData after sheetViews
                    worksheet.AppendChild(sheetData);
                }

                SheetView? sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView == null) {
                    sheetView = new SheetView { WorkbookViewId = 0U };
                    sheetViews.Append(sheetView);
                }

                sheetView.RemoveAllChildren<Pane>();
                sheetView.RemoveAllChildren<Selection>();

                Pane pane = new Pane { State = PaneStateValues.Frozen };
                if (topRows > 0) {
                    pane.VerticalSplit = topRows;  // VerticalSplit = number of rows to freeze
                }
                if (leftCols > 0) {
                    pane.HorizontalSplit = leftCols;  // HorizontalSplit = number of columns to freeze
                }

                pane.TopLeftCell = A1.CellReference(topRows + 1, leftCols + 1);

                if (topRows > 0 && leftCols > 0) {
                    pane.ActivePane = PaneValues.BottomRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else if (topRows > 0) {
                    pane.ActivePane = PaneValues.BottomLeft;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else {
                    pane.ActivePane = PaneValues.TopRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                }

                sheetView.Append(new Selection {
                    ActiveCell = "A1",
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                });

                return true;
            });
        }

        /// <summary>
        /// Shows or hides gridlines on the current sheet (view-level setting).
        /// </summary>
        public void SetGridlinesVisible(bool visible) {
            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                view.ShowGridLines = visible;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Gets whether worksheet row and column headings are displayed in the worksheet view.
        /// </summary>
        public bool RowColumnHeadingsVisible => WorksheetRoot.GetFirstChild<SheetViews>()
            ?.GetFirstChild<SheetView>()
            ?.ShowRowColHeaders
            ?.Value ?? true;

        /// <summary>
        /// Shows or hides worksheet row and column headings in the worksheet view.
        /// </summary>
        public void SetRowColumnHeadingsVisible(bool visible) {
            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                view.ShowRowColHeaders = visible;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Gets whether zero values are displayed in the worksheet view.
        /// </summary>
        public bool ZeroValuesVisible => WorksheetRoot.GetFirstChild<SheetViews>()
            ?.GetFirstChild<SheetView>()
            ?.ShowZeros
            ?.Value ?? true;

        /// <summary>
        /// Shows or hides zero values in the worksheet view.
        /// </summary>
        public void SetZeroValuesVisible(bool visible) {
            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                view.ShowZeros = visible;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Gets whether the worksheet view is displayed from right to left.
        /// </summary>
        public bool RightToLeft => WorksheetRoot.GetFirstChild<SheetViews>()
            ?.GetFirstChild<SheetView>()
            ?.RightToLeft
            ?.Value == true;

        /// <summary>
        /// Sets whether the worksheet view is displayed from right to left.
        /// </summary>
        /// <param name="rightToLeft">True to display the worksheet from right to left; otherwise false.</param>
        public void SetRightToLeft(bool rightToLeft) {
            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                view.RightToLeft = rightToLeft;
                WorksheetRoot.Save();
            });
        }

        /// <summary>
        /// Gets the worksheet view zoom scale percentage, or null when no explicit zoom is configured.
        /// </summary>
        public uint? GetZoomScale() {
            return WorksheetRoot.GetFirstChild<SheetViews>()
                ?.GetFirstChild<SheetView>()
                ?.ZoomScale
                ?.Value;
        }

        /// <summary>
        /// Sets the worksheet view zoom scale percentage.
        /// </summary>
        /// <param name="scale">Zoom percentage from 10 through 400.</param>
        /// <param name="save">Whether to save the worksheet XML immediately.</param>
        public void SetZoomScale(uint scale, bool save = true) {
            if (scale < 10U || scale > 400U) {
                throw new ArgumentOutOfRangeException(nameof(scale), "Worksheet zoom scale must be between 10 and 400 percent.");
            }

            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                view.ZoomScale = scale;
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        internal void SetWorksheetSelection(string activeCell, IReadOnlyList<string> selectedRanges, PaneValues? pane = null, bool save = true) {
            if (string.IsNullOrWhiteSpace(activeCell)) throw new ArgumentNullException(nameof(activeCell));
            if (selectedRanges == null) throw new ArgumentNullException(nameof(selectedRanges));

            string sequenceOfReferences = selectedRanges.Count == 0
                ? activeCell
                : string.Join(" ", selectedRanges.Where(range => !string.IsNullOrWhiteSpace(range)));
            if (string.IsNullOrWhiteSpace(sequenceOfReferences)) {
                sequenceOfReferences = activeCell;
            }

            WriteLock(() => {
                SheetView view = GetOrCreateSheetView();
                foreach (Selection existing in view.Elements<Selection>().Where(selection => SamePane(selection, pane)).ToList()) {
                    existing.Remove();
                }

                var selection = new Selection {
                    ActiveCell = activeCell,
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = sequenceOfReferences }
                };
                if (pane.HasValue) {
                    selection.Pane = pane.Value;
                }

                view.Append(selection);
                if (save) {
                    WorksheetRoot.Save();
                }
            });
        }

        private SheetView GetOrCreateSheetView() {
            Worksheet worksheet = WorksheetRoot;
            SheetViews? sheetViews = worksheet.GetFirstChild<SheetViews>();
            if (sheetViews == null) {
                sheetViews = new SheetViews();
                worksheet.InsertAt(sheetViews, 0);
            }

            SheetView? view = sheetViews.GetFirstChild<SheetView>();
            if (view == null) {
                view = new SheetView { WorkbookViewId = 0U };
                sheetViews.Append(view);
            }

            return view;
        }

        private static bool SamePane(Selection selection, PaneValues? pane) {
            return pane.HasValue
                ? selection.Pane?.Value == pane.Value
                : selection.Pane == null;
        }

        /// <summary>
        /// Configures basic print/page setup for the sheet.
        /// </summary>
        /// <param name="fitToWidth">Number of pages to fit horizontally (1 = fit to one page).</param>
        /// <param name="fitToHeight">Number of pages to fit vertically (0 = unlimited).</param>
        /// <param name="scale">Manual scale (10-400). Ignored if FitToWidth/Height are specified.</param>
        /// <param name="pageOrder">Optional multi-page print order.</param>
        /// <param name="paperSize">Optional known paper size.</param>
        public void SetPageSetup(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null, ExcelPageOrder? pageOrder = null, ExcelPaperSize? paperSize = null) {
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

                ws.Save();
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

                ws.Save();
            });
        }
    }
}
