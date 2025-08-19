using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for autofitting columns and rows.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AutoFitColumnsAndRows() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Long piece of text", autoFitColumns: true, autoFitRows: true);
                sheet.SetCellValue(2, 1, "Second line\nwith newline", autoFitColumns: true, autoFitRows: true);
                sheet.SetCellValue(3, 1, "Line1\nLine2\nLine3", autoFitColumns: true, autoFitRows: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns.Elements<Column>().First();
                Assert.True(column.BestFit.Value);
                Assert.True(column.Width.HasValue && column.Width.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight);
                Assert.True(sheetFormat.DefaultRowHeight > 0);

                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 1);
                Assert.True(row1.CustomHeight);
                Assert.True(row1.Height.HasValue && row1.Height.Value > 0);

                var row3 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 3);
                Assert.True(row3.CustomHeight);
                Assert.True(row3.Height.HasValue && row3.Height.Value > row1.Height.Value);
            }
        }

        [Fact]
        public void Test_AutoFitRows_EmptyRowsRetainDefaultHeight() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.Empty.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Content", autoFitRows: true);
                sheet.SetCellValue(2, 1, " ", autoFitRows: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight);
                Assert.Equal(15.0, sheetFormat.DefaultRowHeight.Value);

                var row2 = wsPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == 2);
                Assert.NotNull(row2);
                Assert.False(row2!.CustomHeight?.Value ?? false);
                Assert.False(row2.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitRows_RemovesCustomHeightWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearRow.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Content", autoFitRows: true);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.SetCellValue(1, 1, string.Empty, autoFitRows: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 1);
                Assert.False(row1.CustomHeight?.Value ?? false);
                Assert.False(row1.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitColumns_RemovesCustomWidthWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Long text", autoFitColumns: true);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.SetCellValue(1, 1, string.Empty, autoFitColumns: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.True(columns == null || !columns.Elements<Column>().Any(c => c.Min == 1 && c.Max == 1));
            }
        }
    }
}
