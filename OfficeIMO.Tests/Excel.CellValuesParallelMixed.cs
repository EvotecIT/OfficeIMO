using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact(Skip = "Flaky in parallel execution on some runtimes")]
        public async Task Test_CellValuesParallel_WithConcurrentCellValue() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesParallelMixed.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                var col1 = Enumerable.Range(1, 500).Select(i => (i, 1, (object)$"R{i}C1"));

                await Task.WhenAll(
                    Task.Run(() => sheet.CellValuesParallel(col1)),
                    Task.Run(() => {
                        for (int i = 1; i <= 500; i++) {
                            sheet.CellValue(i, 2, $"R{i}C2");
                        }
                    })
                );

                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                for (int row = 1; row <= 500; row++) {
                    for (int col = 1; col <= 2; col++) {
                        string expected = $"R{row}C{col}";
                        string cellRef = $"{(char)('A' + col - 1)}{row}";
                        Cell cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRef);
                        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                        int index = int.Parse(cell.CellValue!.Text);
                        Assert.Equal(expected, shared.SharedStringTable!.ElementAt(index).InnerText);
                    }
                }

                Assert.Equal(1000, shared.SharedStringTable!.Count());
            }
        }
    }
}
