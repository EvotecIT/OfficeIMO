using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests concurrent writes to a single worksheet.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_ConcurrentWrites() {
            string filePath = Path.Combine(_directoryWithFiles, "ConcurrentWrites.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                Parallel.For(1, 1001, i => {
                    sheet.CellValue(i, 1, $"Value {i}");
                });
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                for (int i = 1; i <= 1000; i++) {
                    Cell cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == $"A{i}");
                    Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                    int index = int.Parse(cell.CellValue!.Text);
                    Assert.Equal($"Value {i}", shared.SharedStringTable!.ElementAt(index).InnerText);
                }
                Assert.Equal(1000, shared.SharedStringTable!.Count());
            }
        }
    }
}
