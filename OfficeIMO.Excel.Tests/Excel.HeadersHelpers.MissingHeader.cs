using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        [Trait("Category", "ExcelHeaders")]
        public void Excel_SetByHeader_MissingHeader_IsNoOp() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderMissing_SetByHeader.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath)) {
                var sheet = doc.AddWorksheet("Data");
                sheet.Cell(1, 1, "Present");

                sheet.SetByHeader(2, "Missing", "value");

                Assert.False(sheet.TryGetCellText(2, 1, out _));
                doc.Save();
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = pkg.WorkbookPart!.WorksheetParts.First();
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                var row2 = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == 2U);
                var cellA2 = row2?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == "A2");
                Assert.True(cellA2 == null || string.IsNullOrEmpty(cellA2.InnerText));
            }
        }

        [Fact]
        [Trait("Category", "ExcelHeaders")]
        public void Excel_SetByHeader_NullValue_WritesEmptyText() {
            string filePath = Path.Combine(_directoryWithFiles, "Header_SetByHeader_Null.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath)) {
                var sheet = doc.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.SetByHeader(2, "Name", null);

                Assert.True(sheet.TryGetCellText(2, 1, out string text));
                Assert.Equal(string.Empty, text);
                doc.Save();
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = pkg.WorkbookPart!.WorksheetParts.First();
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                var row2 = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == 2U);
                var cellA2 = row2?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == "A2");
                Assert.NotNull(cellA2);
                Assert.Equal(CellValues.String, cellA2!.DataType?.Value);
                Assert.Equal(string.Empty, cellA2.CellValue!.Text);
            }
        }

        [Fact]
        [Trait("Category", "ExcelHeaders")]
        public void Excel_AutoFilterByHeaderEquals_MissingHeader_DoesNothing() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderMissing_AutoFilter.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath)) {
                var sheet = doc.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, "Alice");

                sheet.AutoFilterByHeaderEquals("Missing", new[] { "Alice" });
                doc.Save();
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = pkg.WorkbookPart!.WorksheetParts.First();
                var autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.Null(autoFilter);
            }
        }

        [Fact]
        [Trait("Category", "ExcelHeaders")]
        public void Excel_LinkByHeaderToUrls_MissingHeader_DoesNothing() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderMissing_Links.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath)) {
                var sheet = doc.AddWorksheet("Summary");
                sheet.Cell(1, 1, "Existing");
                sheet.Cell(2, 1, "value");

                sheet.LinkByHeaderToUrls("Missing", rowFrom: 2, rowTo: 2, urlForCellText: _ => "https://example.com");
                doc.Save();
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = pkg.WorkbookPart!.WorksheetParts.First();
                var hyperlinks = wsPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                Assert.Null(hyperlinks);
            }
        }

        [Fact]
        [Trait("Category", "ExcelHeaders")]
        public void Excel_ColumnStyleByHeader_MissingHeader_NoThrow() {
            string filePath = Path.Combine(_directoryWithFiles, "HeaderMissing_Style.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            using (var doc = ExcelDocument.Create(filePath)) {
                var sheet = doc.AddWorksheet("Data");
                sheet.Cell(1, 1, "Existing");
                sheet.CellValue(2, 1, 10);

                sheet.ColumnStyleByHeader("Missing").Bold().Number();
                doc.Save();
            }

            using (var pkg = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = pkg.WorkbookPart!.WorksheetParts.First();
                var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
                var row2 = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == 2U);
                var cellA2 = row2?.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == "A2");
                Assert.NotNull(cellA2);
                Assert.True(string.IsNullOrEmpty(cellA2!.CellValue?.Text) || cellA2.CellValue!.Text == "10");
            }
        }
    }
}
