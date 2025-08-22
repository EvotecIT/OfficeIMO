using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for sequential and parallel cell value operations in Excel sheets.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_CellValues() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValues.xlsx");
            var date = new DateTime(2024, 1, 2);
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Hello");
                sheet.CellValue(2, 1, 123.45);
                sheet.CellValue(3, 1, 678.90m);
                sheet.CellValue(4, 1, date);
                sheet.CellValue(5, 1, true);
                sheet.CellFormula(6, 1, "SUM(A2:A3)");
                sheet.Cell(7, 1, 1.23, "A2+1", "0.00");
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                Assert.NotNull(shared);
                Assert.Equal("Hello", shared.SharedStringTable!.ElementAt(0).InnerText);

                Cell cellString = cells.First(c => c.CellReference == "A1");
                Assert.Equal(CellValues.SharedString, cellString.DataType!.Value);
                Assert.Equal("0", cellString.CellValue!.Text);

                Cell cellDouble = cells.First(c => c.CellReference == "A2");
                Assert.Equal(CellValues.Number, cellDouble.DataType!.Value);
                Assert.Equal("123.45", cellDouble.CellValue!.Text);

                Cell cellDecimal = cells.First(c => c.CellReference == "A3");
                Assert.Equal(CellValues.Number, cellDecimal.DataType!.Value);
                Assert.Equal("678.90", cellDecimal.CellValue!.Text);

                Cell cellDate = cells.First(c => c.CellReference == "A4");
                Assert.Equal(CellValues.Number, cellDate.DataType!.Value);
                Assert.Equal(date.ToOADate().ToString(CultureInfo.InvariantCulture), cellDate.CellValue!.Text);

                Cell cellBool = cells.First(c => c.CellReference == "A5");
                Assert.Equal(CellValues.Boolean, cellBool.DataType!.Value);
                Assert.Equal("1", cellBool.CellValue!.Text);

                Cell cellFormula = cells.First(c => c.CellReference == "A6");
                Assert.NotNull(cellFormula.CellFormula);
                Assert.Equal("SUM(A2:A3)", cellFormula.CellFormula!.Text);

                Cell cellCombined = cells.First(c => c.CellReference == "A7");
                Assert.Equal(CellValues.Number, cellCombined.DataType!.Value);
                Assert.Equal("1.23", cellCombined.CellValue!.Text);
                Assert.NotNull(cellCombined.CellFormula);
                Assert.Equal("A2+1", cellCombined.CellFormula!.Text);
                Assert.NotNull(cellCombined.StyleIndex);
                var styles = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet;
                var formats = styles.NumberingFormats.Elements<NumberingFormat>().ToList();
                Assert.Contains(formats, n => n.FormatCode != null && n.FormatCode.Value == "0.00");
            }
        }

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

