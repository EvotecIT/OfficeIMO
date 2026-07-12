using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using TableColumn = DocumentFormat.OpenXml.Spreadsheet.TableColumn;
using TableExtensionList = DocumentFormat.OpenXml.Spreadsheet.TableExtensionList;
using TableStyleInfo = DocumentFormat.OpenXml.Spreadsheet.TableStyleInfo;
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
        public void Test_InsertDataTableAsTableSetStylePersistsDirectTableVisualOptions() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.DirectVisualStyle.xlsx");
            var table = new DataTable("Sales");
            table.Columns.Add("Region", typeof(string));
            table.Columns.Add("Sales", typeof(int));
            table.Rows.Add("NA", 100);
            table.Rows.Add("EMEA", 200);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.GetOrCreateSheet("Data", SheetNameValidationMode.Sanitize);
                string range = sheet.InsertDataTableAsTable(table, tableName: "Sales", style: TableStyle.TableStyleMedium9);
                sheet.SetTableStyle(
                    range,
                    TableStyle.TableStyleMedium9,
                    showFirstColumn: true,
                    showLastColumn: true,
                    showRowStripes: false,
                    showColumnStripes: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                TableStyleInfo styleInfo = tablePart.Table.TableStyleInfo!;
                Assert.True(styleInfo.ShowFirstColumn!.Value);
                Assert.True(styleInfo.ShowLastColumn!.Value);
                Assert.False(styleInfo.ShowRowStripes!.Value);
                Assert.True(styleInfo.ShowColumnStripes!.Value);
            }
        }

        [Fact]
        public void Test_AddTableConvertsHeaderCellsToTableColumnText() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.NumericHeaders.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Product");
                sheet.CellValue(1, 2, 2024d);
                sheet.CellValue(1, 3, 2025d);
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 100d);
                sheet.CellValue(2, 3, 120d);
                sheet.AddTable("A1:C2", true, "RevenueTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var columns = tablePart.Table.TableColumns!.Elements<TableColumn>().ToList();
                Assert.Equal("2024", columns[1].Name!.Value);
                Assert.Equal("2025", columns[2].Name!.Value);

                Assert.Equal("2024", GetCellText(spreadsheet, wsPart, "B1"));
                Assert.Equal("2025", GetCellText(spreadsheet, wsPart, "C1"));
                var b1 = wsPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                    .First(cell => cell.CellReference?.Value == "B1");
                Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString, b1.DataType!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddTableWithoutHeaderDoesNotWriteAutoFilter() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.NoHeader.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Product");
                sheet.CellValue(1, 2, 2024d);
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 100d);
                sheet.AddTable("A1:B2", false, "RevenueTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                Assert.Equal((uint)0, tablePart.Table.HeaderRowCount!.Value);
                Assert.Null(tablePart.Table.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.AutoFilter>());
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
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

        [Theory]
        [InlineData("B1:A2")]
        [InlineData("A3:A1")]
        [InlineData("B3:A1")]
        public void Test_AddTableInvertedRangeThrows(string range) {
            string safeRange = range.Replace(':', '_');
            string filePath = Path.Combine(_directoryWithFiles, $"Table.Inverted.{safeRange}.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                Assert.Throws<ArgumentException>(() =>
                    sheet.AddTable(range, true, "MyTable", TableStyle.TableStyleMedium9));
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

        [Fact]
        public void Test_SetTableTotalsByNameAndClearTableTotals() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.TotalsByName.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.AddTable("A1:B2", true, "SalesTable", TableStyle.TableStyleMedium9);

                sheet.SetTableTotalsByName("SalesTable", new Dictionary<string, TotalsRowFunctionValues> {
                    ["Name"] = TotalsRowFunctionValues.Count,
                    ["Amount"] = TotalsRowFunctionValues.Sum,
                });

                sheet.ClearTableTotals("SalesTable");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var table = tablePart.Table;

                Assert.False(table.TotalsRowShown?.Value ?? false);
                Assert.Equal((uint)0, table.TotalsRowCount!.Value);
                Assert.All(table.TableColumns!.Elements<TableColumn>(), column => Assert.Null(column.TotalsRowFunction));
            }
        }

        [Fact]
        public void Test_SetTableTotalsByNameRestoresTotalsRowCountAfterClear() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.TotalsRestore.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.AddTable("A1:B2", true, "SalesTable", TableStyle.TableStyleMedium9);

                sheet.SetTableTotalsByName("SalesTable", new Dictionary<string, TotalsRowFunctionValues> {
                    ["Name"] = TotalsRowFunctionValues.Count,
                    ["Amount"] = TotalsRowFunctionValues.Sum,
                });
                sheet.ClearTableTotals("SalesTable");
                sheet.SetTableTotalsByName("SalesTable", new Dictionary<string, TotalsRowFunctionValues> {
                    ["Amount"] = TotalsRowFunctionValues.Sum,
                });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var table = tablePart.Table;
                var columns = table.TableColumns!.Elements<TableColumn>().ToList();

                Assert.True(table.TotalsRowShown?.Value);
                Assert.Equal((uint)1, table.TotalsRowCount!.Value);
                Assert.Null(columns[0].TotalsRowFunction);
                Assert.Equal(TotalsRowFunctionValues.Sum, columns[1].TotalsRowFunction?.Value);
            }
        }

        [Fact]
        public void Test_SetTableStyleUpdatesNamedTableVisualFlags() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.StyleByName.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.AddTable("A1:B2", true, "SalesTable", TableStyle.TableStyleMedium9);

                sheet.SetTableStyle(
                    "SalesTable",
                    TableStyle.TableStyleLight11,
                    showFirstColumn: true,
                    showLastColumn: true,
                    showRowStripes: false,
                    showColumnStripes: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var styleInfo = tablePart.Table.TableStyleInfo!;

                Assert.Equal("TableStyleLight11", styleInfo.Name!.Value);
                Assert.True(styleInfo.ShowFirstColumn?.Value);
                Assert.True(styleInfo.ShowLastColumn?.Value);
                Assert.False(styleInfo.ShowRowStripes?.Value);
                Assert.True(styleInfo.ShowColumnStripes?.Value);
            }
        }

        [Fact]
        public void Test_SetTableStyleInsertsBeforeExtensionList() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.StyleBeforeExtensionList.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Amount");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 2d);
                sheet.AddTable("A1:B2", true, "SalesTable", TableStyle.TableStyleMedium9);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                tablePart.Table.TableStyleInfo!.Remove();
                tablePart.Table.Append(new TableExtensionList());
                tablePart.Table.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document.GetSheet("Data").SetTableStyle("SalesTable", TableStyle.TableStyleLight11);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.First();
                var children = tablePart.Table.ChildElements.ToList();
                int styleIndex = children.FindIndex(child => child is TableStyleInfo);
                int extensionIndex = children.FindIndex(child => child is TableExtensionList);

                Assert.True(styleIndex >= 0);
                Assert.True(extensionIndex >= 0);
                Assert.True(styleIndex < extensionIndex);
            }
        }

        private static string GetCellText(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                .First(item => item.CellReference?.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                && int.TryParse(value, out var sharedStringId)) {
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                return table?.ChildElements[sharedStringId].InnerText ?? string.Empty;
            }

            return value;
        }
    }
}
