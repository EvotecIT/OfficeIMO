using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public class ExcelCellAndRangeTests
    {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            if (cell.DataType?.Value == CellValues.InlineString) {
                return cell.InlineString?.InnerText ?? string.Empty;
            }
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
        public void CanSetSingleCell()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet1 = document.AddSheet("Data");

                    sheet1.Cell(2, 3, "Hello");
                }

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("Hello", GetCellValue(document._spreadSheetDocument, sheetPart, "C2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CellEnforces1BasedIndexing()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    var sheet3 = document.AddSheet("Data");

                    sheet3.Cell(0, 1, "X");
                });
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    var sheet5 = document.AddSheet("Data");

                    sheet5.Cell(1, 0, "X");
                });
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanSetRangeOfValues()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[, ] values =
            {
                {
                    "A",
                    "B"
                },
                {
                    "C",
                    "D"
                }
            };
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                {
                    var sheet7 = document.AddSheet("Data");

                    var range9 = sheet7.Range(ExcelA1.CellReference(1, 1) + ":" + ExcelA1.CellReference(2, 2));
                    range9.SetValues(values);
                }

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath))
            {
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
        public void ClearRangeRemovesPendingDirectCellValuesBeforeSave() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 160; row++) {
                    sheet.CellAt(row, 1).SetValue("secret-" + row.ToString(CultureInfo.InvariantCulture));
                }

                sheet.ClearRange("A1:A160", ExcelClearOptions.Values);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                string worksheetXml = worksheetPart.Worksheet.OuterXml;

                Assert.DoesNotContain("secret-", worksheetXml, StringComparison.Ordinal);
                SharedStringTable? sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStrings != null) {
                    Assert.DoesNotContain(sharedStrings.InnerText, "secret-", StringComparison.Ordinal);
                }
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeEnforces1BasedIndexing()
        {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath))
            {
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    var sheet10 = document.AddSheet("Data");

                    var range12 = sheet10.Range(ExcelA1.CellReference(0, 1) + ":" + ExcelA1.CellReference(1, 1));
                    range12.SetValues(null);
                });
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                {
                    var sheet13 = document.AddSheet("Data");

                    var range15 = sheet13.Range(ExcelA1.CellReference(1, 0) + ":" + ExcelA1.CellReference(1, 1));
                    range15.SetValues(null);
                });
            }

            File.Delete(filePath);
        }

        [Fact]
        public void A1ParsingPreservesSimpleCellAndRangeSemantics() {
            Assert.Equal((12, 28), ExcelA1.ParseCellRef(" ab12 "));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("A"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("A0"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("A1:B2"));
            Assert.Equal((ExcelA1.MaxRows, ExcelA1.MaxColumns), ExcelA1.ParseCellRef("XFD1048576"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("XFE1"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("XFD1048577"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("ZZZZZZZ1"));
            Assert.Equal((0, 0), ExcelA1.ParseCellRef("A2147483648"));

            Assert.True(ExcelA1.TryParseRange(" c10 : a2 ", out int r1, out int c1, out int r2, out int c2));
            Assert.Equal((2, 1, 10, 3), (r1, c1, r2, c2));

            Assert.False(ExcelA1.TryParseRange("A1", out _, out _, out _, out _));
            Assert.False(ExcelA1.TryParseRange("A1:XFE1", out r1, out c1, out r2, out c2));
            Assert.False(ExcelA1.TryParseRange("A1:ZZZZZZZ1", out _, out _, out _, out _));
            Assert.False(ExcelA1.TryParseRange("A1:A2147483648", out _, out _, out _, out _));
            Assert.Equal(28, ExcelA1.ColumnLettersToIndex("a-b1"));
            Assert.Equal(28, ExcelA1.ParseColumnIndexFromCellReference(" AB12 "));
            Assert.Equal(1, ExcelA1.ParseColumnIndexFromCellReference("A0000000001"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReference("A"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReference("A0"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReference("A2147483648"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReference("TOTAL"));
            Assert.Equal(28, ExcelA1.ParseColumnIndexFromCellReferenceFast("ab12"));
            Assert.Equal(8, ExcelA1.ParseColumnIndexFromCellReferenceFast("H2501"));
            Assert.Equal(28, ExcelA1.ParseColumnIndexFromCellReferenceFast(" AB12 "));
            Assert.Equal(1, ExcelA1.ParseColumnIndexFromCellReferenceFast("A0000000001"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceFast("A"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceFast("A0"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceFast("A2147483648"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceFast("AB12X"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceFast("TOTAL"));
            Assert.Equal(1, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("A1"));
            Assert.Equal(8, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("H2501"));
            Assert.Equal(28, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("ab12"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("A0"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("AB12X"));
            Assert.Equal(0, ExcelA1.ParseColumnIndexFromCellReferenceWithKnownRowFast("TOTAL"));
            Assert.True(ExcelA1.TryParseCellReferenceFast("ab12", out int fastRow, out int fastCol));
            Assert.Equal((12, 28), (fastRow, fastCol));
            Assert.True(ExcelA1.TryParseCellReferenceFast(" AB12 ", out fastRow, out fastCol));
            Assert.Equal((12, 28), (fastRow, fastCol));
            Assert.False(ExcelA1.TryParseCellReferenceFast("A0", out _, out _));
            Assert.False(ExcelA1.TryParseCellReferenceFast("AB12X", out _, out _));
            Assert.False(ExcelA1.TryParseCellReferenceFast("A2147483648", out _, out _));
            Assert.True(ExcelA1.TryParseCellReferenceFast("XFD1048576", out _, out _));
            Assert.False(ExcelA1.TryParseCellReferenceFast("XFE1", out _, out _));
            Assert.False(ExcelA1.TryParseCellReferenceFast("XFD1048577", out _, out _));
            Assert.Equal("A", ExcelA1.ColumnIndexToLetters(0));
            Assert.Equal("A", ExcelA1.ColumnIndexToLetters(1));
            Assert.Equal("Z", ExcelA1.ColumnIndexToLetters(26));
            Assert.Equal("AA", ExcelA1.ColumnIndexToLetters(27));
            Assert.Equal("AB", ExcelA1.ColumnIndexToLetters(28));
            Assert.Equal("XFD", ExcelA1.ColumnIndexToLetters(16384));
            Assert.Equal("AB12", ExcelA1.CellReference(12, 28));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelA1.CellReference(0, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => ExcelA1.CellReference(1, 0));
        }

        [Fact]
        public void CellAndRangeStylePresetsApplyFormatsAndStatusColors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Styles");
                sheet.CellAt(1, 1).SetValue("Amount").HeaderStyle();
                sheet.CellAt(2, 1).SetValue(123.45).Currency(culture: CultureInfo.GetCultureInfo("en-US")).Success();
                sheet.Range("B2:B3").Percent(1).Warning();
                sheet.CellAt(2, 2).SetValue(0.42);
                sheet.CellAt(3, 2).SetValue(0.67);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                string stylesXml = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet.OuterXml;

                Assert.Contains("$", stylesXml, StringComparison.Ordinal);
                Assert.Contains("0.0%", stylesXml, StringComparison.Ordinal);
                Assert.Contains("E7F6E7", stylesXml, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("FFF4CC", stylesXml, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("D9EAF7", stylesXml, StringComparison.OrdinalIgnoreCase);
            }

            File.Delete(filePath);
        }
    }
}