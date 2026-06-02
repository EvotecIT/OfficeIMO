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
                var worksheet = WorksheetRoot;
                var sheetViews = worksheet.GetFirstChild<SheetViews>();
                if (sheetViews == null) {
                    sheetViews = new SheetViews();
                    worksheet.InsertAt(sheetViews, 0);
                }
                var view = sheetViews.GetFirstChild<SheetView>();
                if (view == null) {
                    view = new SheetView { WorkbookViewId = 0U };
                    sheetViews.Append(view);
                }
                view.ShowGridLines = visible;
                worksheet.Save();
            });
        }

        /// <summary>
        /// Configures basic print/page setup for the sheet.
        /// </summary>
        /// <param name="fitToWidth">Number of pages to fit horizontally (1 = fit to one page).</param>
        /// <param name="fitToHeight">Number of pages to fit vertically (0 = unlimited).</param>
        /// <param name="scale">Manual scale (10-400). Ignored if FitToWidth/Height are specified.</param>
        public void SetPageSetup(uint? fitToWidth = null, uint? fitToHeight = null, uint? scale = null) {
            WriteLock(() => {
                var ws = WorksheetRoot;
                var pageSetup = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageSetup>();
                if (pageSetup == null) {
                    pageSetup = new DocumentFormat.OpenXml.Spreadsheet.PageSetup();
                    // Insert after PageMargins when present, else at end
                    var margins = ws.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.PageMargins>();
                    if (margins != null) ws.InsertAfter(pageSetup, margins); else ws.Append(pageSetup);
                }

                if (fitToWidth != null) pageSetup.FitToWidth = fitToWidth.Value;
                if (fitToHeight != null) pageSetup.FitToHeight = fitToHeight.Value;
                if (scale != null) pageSetup.Scale = scale.Value;

                ws.Save();
            });
        }
    }
}
