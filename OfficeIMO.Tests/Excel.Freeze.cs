using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for freezing rows and columns.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_FreezeTopRows() {
            string filePath = Path.Combine(_directoryWithFiles, "FreezeTopRows.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s.Freeze(topRows: 1))
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                SheetView sheetView = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
                Pane pane = sheetView?.GetFirstChild<Pane>();
                Assert.NotNull(pane);
                Assert.Equal(1D, pane!.HorizontalSplit?.Value);
                Assert.Null(pane.VerticalSplit);
                Assert.Equal(PaneValues.BottomLeft, pane.ActivePane?.Value);
                Assert.Equal("A2", pane.TopLeftCell?.Value);

                Selection selection = sheetView!.GetFirstChild<Selection>();
                Assert.NotNull(selection);
                Assert.Equal(PaneValues.BottomLeft, selection!.Pane?.Value);
                Assert.Equal("A2", selection.ActiveCell?.Value);
                Assert.Equal("A2", selection.SequenceOfReferences?.InnerText);
            }
        }

        [Fact]
        public void Test_FreezeLeftColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "FreezeLeftColumns.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s.Freeze(leftCols: 2))
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                SheetView sheetView = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
                Pane pane = sheetView?.GetFirstChild<Pane>();
                Assert.NotNull(pane);
                Assert.Equal(2D, pane!.VerticalSplit?.Value);
                Assert.Null(pane.HorizontalSplit);
                Assert.Equal(PaneValues.TopRight, pane.ActivePane?.Value);
                Assert.Equal("C1", pane.TopLeftCell?.Value);

                Selection selection = sheetView!.GetFirstChild<Selection>();
                Assert.NotNull(selection);
                Assert.Equal(PaneValues.TopRight, selection!.Pane?.Value);
                Assert.Equal("C1", selection.ActiveCell?.Value);
                Assert.Equal("C1", selection.SequenceOfReferences?.InnerText);
            }
        }

        [Fact]
        public void Test_FreezeTopRowsAndLeftColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "FreezeTopRowsAndLeftColumns.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => s.Freeze(topRows: 1, leftCols: 1))
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                SheetView sheetView = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
                Pane pane = sheetView?.GetFirstChild<Pane>();
                Assert.NotNull(pane);
                Assert.Equal(1D, pane!.HorizontalSplit?.Value);
                Assert.Equal(1D, pane.VerticalSplit?.Value);
                Assert.Equal(PaneValues.BottomRight, pane.ActivePane?.Value);
                Assert.Equal("B2", pane.TopLeftCell?.Value);

                Selection[] selections = sheetView!.Elements<Selection>().ToArray();
                Assert.Equal(3, selections.Length);
                foreach (Selection selection in selections) {
                    Assert.Equal(pane.TopLeftCell?.Value, selection.ActiveCell?.Value);
                    Assert.Equal(pane.TopLeftCell?.Value, selection.SequenceOfReferences?.InnerText);
                }
            }
        }
    }
}

