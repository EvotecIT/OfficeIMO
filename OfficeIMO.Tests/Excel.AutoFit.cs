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
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns!.Elements<Column>().First();
                Assert.True(column.BestFit?.Value ?? false);
                Assert.True(column.Width is { } w && w.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight!.Value);
                Assert.True(sheetFormat.DefaultRowHeight!.Value > 0);

                var row = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 1)!;
                Assert.True(row.CustomHeight!.Value);
                Assert.True(row.Height is { } h && h.Value > 0);
            }
        }
    }
}
