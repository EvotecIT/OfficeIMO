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
                Pane pane = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>();
                Assert.NotNull(pane);
                Assert.Equal(1D, pane!.HorizontalSplit?.Value);
                Assert.Null(pane.VerticalSplit);
                Assert.Equal(PaneValues.BottomLeft, pane.ActivePane?.Value);
                Assert.Equal("A2", pane.TopLeftCell?.Value);
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
                Pane pane = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>()?.GetFirstChild<Pane>();
                Assert.NotNull(pane);
                Assert.Equal(2D, pane!.VerticalSplit?.Value);
                Assert.Null(pane.HorizontalSplit);
                Assert.Equal(PaneValues.TopRight, pane.ActivePane?.Value);
                Assert.Equal("C1", pane.TopLeftCell?.Value);
            }
        }
    }
}

