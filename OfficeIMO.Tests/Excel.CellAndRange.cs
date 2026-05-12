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
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int id)) {
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
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Hello", GetCellValue(document._spreadSheetDocument, sheetPart, "C2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CellEnforces1BasedIndexing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Sheet("Data", s => s.Cell(0, 1, "X")));
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Sheet("Data", s => s.Cell(1, 0, "X")));
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
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
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
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Sheet("Data", s => s.Range(1, 0, 1, 1, null)));
            }
            File.Delete(filePath);
        }

        [Fact]
        public void A1ParsingPreservesSimpleCellAndRangeSemantics() {
            Assert.Equal((12, 28), A1.ParseCellRef(" ab12 "));
            Assert.Equal((0, 0), A1.ParseCellRef("A"));
            Assert.Equal((0, 0), A1.ParseCellRef("A0"));
            Assert.Equal((0, 0), A1.ParseCellRef("A1:B2"));
            Assert.Equal((A1.MaxRows, A1.MaxColumns), A1.ParseCellRef("XFD1048576"));
            Assert.Equal((1, A1.MaxColumns + 1), A1.ParseCellRef("XFE1"));
            Assert.Equal((A1.MaxRows + 1, A1.MaxColumns), A1.ParseCellRef("XFD1048577"));
            Assert.Equal((0, 0), A1.ParseCellRef("ZZZZZZZ1"));
            Assert.Equal((0, 0), A1.ParseCellRef("A2147483648"));

            Assert.True(A1.TryParseRange(" c10 : a2 ", out int r1, out int c1, out int r2, out int c2));
            Assert.Equal((2, 1, 10, 3), (r1, c1, r2, c2));

            Assert.False(A1.TryParseRange("A1", out _, out _, out _, out _));
            Assert.True(A1.TryParseRange("A1:XFE1", out r1, out c1, out r2, out c2));
            Assert.Equal((1, 1, 1, A1.MaxColumns + 1), (r1, c1, r2, c2));
            Assert.False(A1.TryParseRange("A1:ZZZZZZZ1", out _, out _, out _, out _));
            Assert.False(A1.TryParseRange("A1:A2147483648", out _, out _, out _, out _));
            Assert.Equal(28, A1.ColumnLettersToIndex("a-b1"));
            Assert.Equal("A", A1.ColumnIndexToLetters(0));
            Assert.Equal("A", A1.ColumnIndexToLetters(1));
            Assert.Equal("Z", A1.ColumnIndexToLetters(26));
            Assert.Equal("AA", A1.ColumnIndexToLetters(27));
            Assert.Equal("AB", A1.ColumnIndexToLetters(28));
            Assert.Equal("XFD", A1.ColumnIndexToLetters(16384));
            Assert.Equal("AB12", A1.CellReference(12, 28));
            Assert.Throws<ArgumentOutOfRangeException>(() => A1.CellReference(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => A1.CellReference(1, 0));
        }
    }
}
