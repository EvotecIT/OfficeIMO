using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
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

                Selection[] selections = sheetView!.Elements<Selection>().ToArray();
                Assert.Equal(2, selections.Length);
                Selection bottomLeft = selections.Single(s => s.Pane?.Value == PaneValues.BottomLeft);
                Assert.Equal("A2", bottomLeft.ActiveCell?.Value);
                Assert.Equal("A2", bottomLeft.SequenceOfReferences?.InnerText);
                Selection topLeft = selections.Single(s => s.Pane == null);
                Assert.Equal("A1", topLeft.ActiveCell?.Value);
                Assert.Equal("A1", topLeft.SequenceOfReferences?.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(wsPart.Worksheet).ToList();
                Assert.Empty(errors);
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

                Selection[] selections = sheetView!.Elements<Selection>().ToArray();
                Assert.Equal(2, selections.Length);
                Selection topRight = selections.Single(s => s.Pane?.Value == PaneValues.TopRight);
                Assert.Equal("C1", topRight.ActiveCell?.Value);
                Assert.Equal("C1", topRight.SequenceOfReferences?.InnerText);
                Selection topLeft = selections.Single(s => s.Pane == null);
                Assert.Equal("A1", topLeft.ActiveCell?.Value);
                Assert.Equal("A1", topLeft.SequenceOfReferences?.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(wsPart.Worksheet).ToList();
                Assert.Empty(errors);
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
                Assert.Equal(4, selections.Length);
                Selection topRight = selections.Single(s => s.Pane?.Value == PaneValues.TopRight);
                Selection bottomLeft = selections.Single(s => s.Pane?.Value == PaneValues.BottomLeft);
                Selection bottomRight = selections.Single(s => s.Pane?.Value == PaneValues.BottomRight);
                Selection topLeft = selections.Single(s => s.Pane == null);
                foreach (Selection sel in new[] { topRight, bottomLeft, bottomRight }) {
                    Assert.Equal(pane.TopLeftCell?.Value, sel.ActiveCell?.Value);
                    Assert.Equal(pane.TopLeftCell?.Value, sel.SequenceOfReferences?.InnerText);
                }
                Assert.Equal("A1", topLeft.ActiveCell?.Value);
                Assert.Equal("A1", topLeft.SequenceOfReferences?.InnerText);

                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(wsPart.Worksheet).ToList();
                Assert.Empty(errors);
            }
        }

        [Fact]
        public void Test_UnfreezeRemovesSheetViews() {
            string filePath = Path.Combine(_directoryWithFiles, "FreezeUnfreeze.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.AsFluent()
                    .Sheet("Data", s => {
                        s.Freeze(topRows: 1, leftCols: 1);
                        s.Freeze();
                    })
                    .End()
                    .Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.GetFirstChild<SheetViews>());

                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(wsPart.Worksheet).ToList();
                Assert.Empty(errors);
            }
        }
    }
}

