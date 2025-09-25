using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using TableColumn = DocumentFormat.OpenXml.Spreadsheet.TableColumn;
using TotalsRowFunctionValues = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddTableWithStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                Assert.Equal("A1:B2", tablePart.Table.Reference!.Value);
                Assert.Equal("MyTable", tablePart.Table.Name);
                Assert.Equal("TableStyleMedium9", tablePart.Table.TableStyleInfo!.Name!.Value);
            }
        }

        [Fact]
        public void Test_AddTablePopulatesMissingCells() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.MissingCells.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>();
                Assert.Contains(cells, c => c.CellReference == "A2");
                Assert.Contains(cells, c => c.CellReference == "B2");
            }
        }

        [Fact]
        public async Task Test_AddTableConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Concurrent.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                for (int i = 0; i < 5; i++) {
                    int rowStart = 1 + i * 3;
                    sheet.CellValue(rowStart, 1, "Name");
                    sheet.CellValue(rowStart, 2, "Value");
                    sheet.CellValue(rowStart + 1, 1, "A");
                    sheet.CellValue(rowStart + 1, 2, 1d);
                    sheet.CellValue(rowStart + 2, 1, "B");
                    sheet.CellValue(rowStart + 2, 2, 2d);
                }

                var tasks = Enumerable.Range(0, 5)
                    .Select(i => Task.Run(() => sheet.AddTable($"A{1 + i * 3}:B{3 + i * 3}", true, $"MyTable{i}", TableStyle.TableStyleMedium9)))
                    .ToArray();
                await Task.WhenAll(tasks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Equal(5, wsPart.TableDefinitionParts.Count());
            }
        }

        [Fact]
        public void Test_AddTableOverlappingRangeThrows() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Overlap.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.AddTable("A1:B3", true, "Table1", TableStyle.TableStyleMedium9);

                Assert.Throws<InvalidOperationException>(() =>
                    sheet.AddTable("B2:C4", true, "Table2", TableStyle.TableStyleMedium9));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Single(wsPart.TableDefinitionParts);
                  Assert.Equal("A1:B3", wsPart.TableDefinitionParts.First().Table.Reference!.Value);
            }
        }

        [Fact]
        public void Test_SetTableTotalsMatchesHeadersCaseInsensitive() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.TotalsCaseInsensitive.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.AddTable("A1:B2", true, "MyTable", TableStyle.TableStyleMedium9);
                sheet.SetTableTotals("A1:B2", new Dictionary<string, TotalsRowFunctionValues> {
                    ["name"] = TotalsRowFunctionValues.Count,
                    ["AMOUNT"] = TotalsRowFunctionValues.Sum,
                });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var table = tablePart.Table;

                Assert.True(table.TotalsRowShown?.Value);

                var columns = table.TableColumns!.Elements<TableColumn>().ToList();
                Assert.Equal(TotalsRowFunctionValues.Count, columns[0].TotalsRowFunction?.Value);
                Assert.Equal(TotalsRowFunctionValues.Sum, columns[1].TotalsRowFunction?.Value);
            }
        }

        [Fact]
        public void Test_SetTableTotalsMatchesHeadersWithMixedCasing() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.TotalsMixedCase.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "nAmE");
                sheet.CellValue(1, 2, "AMOUNT");
                sheet.CellValue(1, 3, "balance");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.CellValue(2, 3, 3d);
                sheet.AddTable("A1:C2", true, "MyTable", TableStyle.TableStyleMedium9);
                sheet.SetTableTotals("A1:C2", new Dictionary<string, TotalsRowFunctionValues> {
                    ["NAME"] = TotalsRowFunctionValues.Count,
                    ["amount"] = TotalsRowFunctionValues.Sum,
                    ["BaLaNcE"] = TotalsRowFunctionValues.Sum,
                });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var table = tablePart.Table;

                Assert.True(table.TotalsRowShown?.Value);

                var columns = table.TableColumns!.Elements<TableColumn>().ToList();
                Assert.Equal(TotalsRowFunctionValues.Count, columns[0].TotalsRowFunction?.Value);
                Assert.Equal(TotalsRowFunctionValues.Sum, columns[1].TotalsRowFunction?.Value);
                Assert.Equal(TotalsRowFunctionValues.Sum, columns[2].TotalsRowFunction?.Value);
            }
        }
    }
}
