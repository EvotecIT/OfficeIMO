using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_EnumerateCells_ReturnsCorrectCoordinates() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateCellsCoordinates.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "B2");
                    sheet.CellValue(3, 4, "D3");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateCells().ToList();

                Assert.Contains(cells, c => c.Row == 2 && c.Column == 2 && Equals(c.Value, "B2"));
                Assert.Contains(cells, c => c.Row == 3 && c.Column == 4 && Equals(c.Value, "D3"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_EnumerateRange_FiltersUsingActualColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateRangeColumns.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(2, 2, "Inside");
                    sheet.CellValue(2, 4, "Outside");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateRange("B1:C3").ToList();

                var onlyCell = Assert.Single(cells);
                Assert.Equal(2, onlyCell.Row);
                Assert.Equal(2, onlyCell.Column);
                Assert.Equal("Inside", onlyCell.Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_EnumerateRange_DoesNotStopAtOutOfOrderRowsBeyondRange() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderEnumerateRangeOutOfOrderRows.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(2, 1, "InRange");
                    sheet.CellValue(5, 1, "Outside");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
                    var row2 = sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 2U);
                    row2.Remove();
                    sheetData.Append(row2);
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var cells = reader.GetSheet("Data").EnumerateRange("A1:A2").ToList();

                Assert.Contains(cells, c => c.Row == 1 && c.Column == 1 && Equals(c.Value, "Header"));
                Assert.Contains(cells, c => c.Row == 2 && c.Column == 1 && Equals(c.Value, "InRange"));
                Assert.DoesNotContain(cells, c => c.Row == 5);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_FormulaText_IsReturnedWhenCachedResultsAreDisabled() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderFormulaText.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, 2);
                    sheet.CellValue(2, 1, 3);
                    sheet.CellFormula(3, 1, "=SUM(A1:A2)");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath, new ExcelReadOptions { UseCachedFormulaResult = false });
                var cell = reader.GetSheet("data").EnumerateCells().Single(c => c.Row == 3 && c.Column == 1);

                Assert.Equal("SUM(A1:A2)", cell.Value);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
