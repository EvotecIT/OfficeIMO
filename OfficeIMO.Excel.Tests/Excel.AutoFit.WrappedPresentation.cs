using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("Mean distance from Sun (au)", 24.070026369762672, true)]
        [InlineData("Extensive system", 13.741869323176092, false)]
        public void Test_AutoFitRows_LeavesHeightForNativeWrappedLabels(string text, double width, bool bold) {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.NativeWrapped." + width.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.SetColumnWidth(1, width);
                sheet.CellValue(1, 1, text);
                sheet.CellAt(1, 1).SetBold(bold);
                sheet.CellValue(1, 2, "Neighbor");
                sheet.CellWrapText(1, 1);
                sheet.CellValue(2, 1, "Short control");
                sheet.CellWrapText(2, 1);
                sheet.AutoFitRows();
                document.Save();
            }

            using (var reopened = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Equal(text, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
                Assert.Equal("Neighbor", reopened.Sheets[0].CellAt(1, 2).GetValue<string>());
                Assert.Equal(width, reopened.Sheets[0].GetColumnDefinitions().First(c => c.StartIndex == 1).Width);
            }
            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var rows = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Row>().ToList();
                // These labels wrap to two native lines at the authored widths. A single-line
                // explicit height clips the first line even though the saved values are intact.
                Assert.True(rows.Single(r => r.RowIndex?.Value == 1).Height!.Value >= 30D);
                Assert.True(rows.Single(r => r.RowIndex?.Value == 2).Height!.Value < 30D);
            }
        }
    }
}
