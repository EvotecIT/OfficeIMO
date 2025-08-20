using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelCellAndRangeTests {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                if (int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        [Fact]
        public void CanSetSingleCell() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Cell(2, 3, "Hello"));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                var sheetPart = document._spreadSheetDocument.WorkbookPart.WorksheetParts.First();
                Assert.Equal("Hello", GetCellValue(document._spreadSheetDocument, sheetPart, "C2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CellEnforces1BasedIndexing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Sheet("Data", s => s.Cell(0, 1, "X")));
            }
            File.Delete(filePath);
        }

        [Fact]
        public void CanSetRangeOfValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[,] values = {
                { "A", "B" },
                { "C", "D" }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Range(1, 1, 2, 2, values));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                var sheetPart = document._spreadSheetDocument.WorkbookPart.WorksheetParts.First();
                Assert.Equal("A", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("D", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeEnforces1BasedIndexing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Sheet("Data", s => s.Range(0, 1, 1, 1, null)));
            }
            File.Delete(filePath);
        }
    }
}
