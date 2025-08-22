using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_CellValuesParallel() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesParallel.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                var col1 = Enumerable.Range(1, 1000).Select(i => (i, 1, (object)$"R{i}C1"));
                var col2 = Enumerable.Range(1, 1000).Select(i => (i, 2, (object)$"R{i}C2"));
                var col3 = Enumerable.Range(1, 1000).Select(i => (i, 3, (object)$"R{i}C3"));
                var col4 = Enumerable.Range(1, 1000).Select(i => (i, 4, (object)$"R{i}C4"));
                var col5 = Enumerable.Range(1, 1000).Select(i => (i, 5, (object)$"R{i}C5"));
                var col6 = Enumerable.Range(1, 1000).Select(i => (i, 6, (object)$"R{i}C6"));
                var col7 = Enumerable.Range(1, 1000).Select(i => (i, 7, (object)$"R{i}C7"));
                var col8 = Enumerable.Range(1, 1000).Select(i => (i, 8, (object)$"R{i}C8"));

                await Task.WhenAll(
                    Task.Run(() => sheet.CellValuesParallel(col1)),
                    Task.Run(() => sheet.CellValuesParallel(col2)),
                    Task.Run(() => sheet.CellValuesParallel(col3)),
                    Task.Run(() => sheet.CellValuesParallel(col4)),
                    Task.Run(() => sheet.CellValuesParallel(col5)),
                    Task.Run(() => sheet.CellValuesParallel(col6)),
                    Task.Run(() => sheet.CellValuesParallel(col7)),
                    Task.Run(() => sheet.CellValuesParallel(col8))
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

                for (int row = 1; row <= 1000; row++) {
                    for (int col = 1; col <= 8; col++) {
                        string expected = $"R{row}C{col}";
                        string cellRef = $"{(char)('A' + col - 1)}{row}";
                        Cell cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRef);
                        Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                        int index = int.Parse(cell.CellValue!.Text);
                        Assert.Equal(expected, shared.SharedStringTable!.ElementAt(index).InnerText);
                    }
                }
                OpenXmlValidator validator = new OpenXmlValidator();
                Assert.Empty(validator.Validate(spreadsheet));

                Assert.Equal(8000, shared.SharedStringTable!.Count());
            }
        }
    }
}

