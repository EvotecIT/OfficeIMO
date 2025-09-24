using System;
using System.Collections.Generic;
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
        private enum DummyEnum { Alpha, Beta }

        [Fact]
        public async Task Test_SetCellValuesParallelStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "SetCellValuesParallelStrings.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                var col1 = Enumerable.Range(1, 500).Select(i => (i, 1, (object)$"R{i}C1"));
                var col2 = Enumerable.Range(1, 500).Select(i => (i, 2, (object)$"R{i}C2"));
                var col3 = Enumerable.Range(1, 500).Select(i => (i, 3, (object)$"R{i}C3"));
                var col4 = Enumerable.Range(1, 500).Select(i => (i, 4, (object)$"R{i}C4"));

                await Task.WhenAll(
                    Task.Run(() => sheet.SetCellValues(col1, ExecutionMode.Parallel)),
                    Task.Run(() => sheet.SetCellValues(col2, ExecutionMode.Parallel)),
                    Task.Run(() => sheet.SetCellValues(col3, ExecutionMode.Parallel)),
                    Task.Run(() => sheet.SetCellValues(col4, ExecutionMode.Parallel))
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
                    for (int col = 1; col <= 4; col++) {
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

                Assert.Equal(2000, shared.SharedStringTable!.Count());
            }
        }

        [Fact]
        public void Test_SetCellValuesParallelMixedTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "SetCellValuesParallelMixedTypes.xlsx");
            Guid guid = Guid.NewGuid();
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var uri = new Uri("https://example.com");
                var cells = new (int, int, object)[] {
                    (1, 1, (object)guid),
                    (1, 2, (object)DummyEnum.Beta),
                    (1, 3, (object)'Z'),
                    (1, 4, (object)DBNull.Value),
                    (1, 5, (object)uri)
                };
                sheet.SetCellValues(cells, ExecutionMode.Parallel);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;

                string cellRefA1 = "A1";
                Cell a1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRefA1);
                Assert.Equal(CellValues.SharedString, a1.DataType!.Value);
                int idx = int.Parse(a1.CellValue!.Text);
                Assert.Equal(guid.ToString(), shared.SharedStringTable!.ElementAt(idx).InnerText);

                Cell b1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "B1");
                Assert.Equal(CellValues.SharedString, b1.DataType!.Value);
                idx = int.Parse(b1.CellValue!.Text);
                Assert.Equal("Beta", shared.SharedStringTable!.ElementAt(idx).InnerText);

                Cell c1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "C1");
                Assert.Equal(CellValues.SharedString, c1.DataType!.Value);
                idx = int.Parse(c1.CellValue!.Text);
                Assert.Equal("Z", shared.SharedStringTable!.ElementAt(idx).InnerText);

                Cell d1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "D1");
                Assert.Equal(CellValues.String, d1.DataType!.Value);
                Assert.True(string.IsNullOrEmpty(d1.CellValue?.Text));

                Cell e1 = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "E1");
                Assert.Equal(CellValues.SharedString, e1.DataType!.Value);
                idx = int.Parse(e1.CellValue!.Text);
                Assert.Equal("https://example.com/", shared.SharedStringTable!.ElementAt(idx).InnerText);
            }
        }

        [Fact]
        public void Test_SetCellValuesParallelSanitizesIllegalControlCharacters() {
            string filePath = Path.Combine(_directoryWithFiles, "SetCellValuesParallelSanitizeControls.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var inputs = new (int Row, int Column, object Value)[] {
                    (1, 1, (object)"Value\u0001With\u0002Controls"),
                    (1, 2, (object)"ValueWithControls"),
                    (1, 3, (object)"Value\u0003With\u0004Controls"),
                    (2, 1, (object)"\u0005Leading"),
                    (2, 2, (object)"Trailing\u0006"),
                    (2, 3, (object)"Tab\tAllowed"),
                };

                sheet.SetCellValues(inputs, ExecutionMode.Parallel);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                SharedStringTable sharedTable = shared.SharedStringTable!;

                OpenXmlValidator validator = new OpenXmlValidator();
                Assert.Empty(validator.Validate(spreadsheet));

                var expected = new Dictionary<string, string> {
                    ["A1"] = "ValueWithControls",
                    ["B1"] = "ValueWithControls",
                    ["C1"] = "ValueWithControls",
                    ["A2"] = "Leading",
                    ["B2"] = "Trailing",
                    ["C2"] = "Tab\tAllowed",
                };

                var indices = new Dictionary<string, int>();

                foreach (var kvp in expected) {
                    Cell cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == kvp.Key);
                    Assert.Equal(CellValues.SharedString, cell.DataType!.Value);
                    int index = int.Parse(cell.CellValue!.Text);
                    Assert.Equal(kvp.Value, sharedTable.ElementAt(index).InnerText);
                    indices[kvp.Key] = index;
                }

                Assert.Equal(indices["A1"], indices["B1"]);
                Assert.Equal(indices["A1"], indices["C1"]);
                Assert.Equal(4, sharedTable.Count());
            }
        }

        [Fact]
        public void Test_CellValueThrowsOnTooLongString() {
            string filePath = Path.Combine(_directoryWithFiles, "TooLongString.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                string longText = new string('a', 32768);
                Assert.Throws<ArgumentException>(() => sheet.CellValue(1, 1, longText));
            }
        }
    }
}

