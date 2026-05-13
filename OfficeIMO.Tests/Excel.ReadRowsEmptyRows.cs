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
        public void ReadRows_ReturnsNullForRowsWithoutCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReadRowsEmptyRows.xlsx");
            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(3, 1, "Value");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadRows("A1:A3").ToList();

                Assert.Equal(3, rows.Count);
                Assert.NotNull(rows[0]);
                Assert.Equal("Header", rows[0]![0]);
                Assert.Null(rows[1]);
                Assert.NotNull(rows[2]);
                Assert.Equal("Value", rows[2]![0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ReadRows_ReturnsBlankArrayForRowsWithBlankCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReadRowsBlankCells.xlsx");
            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(3, 1, "Value");
                    document.Save();
                }

                using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet!.GetFirstChild<SheetData>()!;
                    var row = new Row { RowIndex = 2U };
                    row.Append(new Cell { CellReference = "A2", StyleIndex = 0U });
                    sheetData.InsertBefore(row, sheetData.Elements<Row>().First(r => r.RowIndex?.Value == 3U));
                    worksheetPart.Worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var rows = reader.GetSheet("Data").ReadRows("A1:A3").ToList();

                Assert.Equal(3, rows.Count);
                Assert.NotNull(rows[1]);
                Assert.Null(rows[1]![0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ReadRowsAs_ThrowsForRowsWithoutCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReadRowsAsEmptyRows.xlsx");
            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddWorkSheet("Data");
                    sheet.CellValue(1, 1, "Header");
                    sheet.CellValue(3, 1, "Value");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                var sheetReader = reader.GetSheet("Data");

                var ex = Assert.Throws<InvalidOperationException>(() => sheetReader.ReadRowsAs<string>("A1:A3").ToList());
                Assert.Contains("Row 2", ex.Message);
                Assert.Contains("contains no cells", ex.Message);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
