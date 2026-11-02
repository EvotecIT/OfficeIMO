using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using SpreadsheetTable = DocumentFormat.OpenXml.Spreadsheet.Table;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public sealed class TotalsOrder {
            public string Customer { get; set; } = string.Empty;
            public DateTime Date { get; set; }
            public double Amount { get; set; }
        }

        private static ExcelSheet CreateOrdersTable(ExcelDocument document, int orders) {
            ExcelSheet sheet = document.AddSheet("Orders");
            sheet.Cell(1, 1).SetValue("Customer");
            sheet.Cell(1, 2).SetValue("Date");
            sheet.Cell(1, 3).SetValue("Amount");
            for (int index = 0; index < orders; index++) {
                sheet.Cell(index + 2, 1).SetValue("Customer " + (index + 1));
                sheet.Cell(index + 2, 2).SetValue(new DateTime(2026, 1, 1).AddDays(index));
                sheet.Cell(index + 2, 3).SetValue(10d * (index + 1));
            }

            sheet.AddTable("A1:C" + (orders + 1), true, "Orders", ExcelTableStyle.TableStyleMedium2);
            return sheet;
        }

        [Fact]
        public void SetTableTotals_AddsTotalsRowBelowDataWithSubtotalFormula() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = CreateOrdersTable(document, 3);
                sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                    ["Amount"] = ExcelTableTotalsFunction.Sum
                });

                // The last order stays a data row.
                Assert.Equal(new[] { "Customer 3" }, sheet.RowsAs<TotalsOrder>().Skip(2).Select(order => order.Customer).ToArray());
                Assert.Equal(3, sheet.RowsAs<TotalsOrder>().Count());
                document.Save();
            }

            stream.Position = 0;
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
            WorksheetPart worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First();
            SpreadsheetTable table = worksheet.TableDefinitionParts.First().Table!;
            Assert.Equal("A1:C5", table.Reference!.Value);
            Assert.Equal(1U, table.TotalsRowCount!.Value);
            Cell[] totals = worksheet.Worksheet.Descendants<Cell>().Where(cell => cell.CellReference!.Value!.EndsWith("5", StringComparison.Ordinal)).ToArray();
            Cell total = Assert.Single(totals, cell => cell.CellReference!.Value == "C5");
            Assert.Equal("SUBTOTAL(109,Orders[Amount])", total.CellFormula!.Text);
            Assert.Equal("60", total.CellValue!.Text);
            Assert.Contains(totals, cell => cell.CellReference!.Value == "A5");
        }

        [Fact]
        public void SetTableTotals_ThrowsWhenTheRowBelowTheTableIsOccupied() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateOrdersTable(document, 2);
            sheet.Cell(4, 2).SetValue("Notes below the table");

            Assert.Throws<InvalidOperationException>(() => sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                ["Amount"] = ExcelTableTotalsFunction.Sum
            }));
        }

        [Fact]
        public void SetTableTotals_ReusesAReservedBlankLastRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Orders");
            sheet.Cell(1, 1).SetValue("Customer");
            sheet.Cell(1, 2).SetValue("Amount");
            sheet.Cell(2, 1).SetValue("Contoso");
            sheet.Cell(2, 2).SetValue(5d);
            sheet.Cell(3, 1).SetValue("Fabrikam");
            sheet.Cell(3, 2).SetValue(7d);
            sheet.AddTable("A1:B4", true, "Orders", ExcelTableStyle.TableStyleMedium2);

            sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                ["Amount"] = ExcelTableTotalsFunction.Sum
            });

            Assert.Equal("A1:B4", document.GetTables().Single().Range);
        }

        [Fact]
        public void ClearTableTotals_RemovesTheTotalsRowAndRestoresTheRange() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateOrdersTable(document, 2);
            var totals = new Dictionary<string, ExcelTableTotalsFunction> { ["Amount"] = ExcelTableTotalsFunction.Sum };
            sheet.SetTableTotalsByName("Orders", totals);
            Assert.Equal("A1:C4", document.GetTables().Single().Range);

            sheet.ClearTableTotals("Orders");
            Assert.Equal("A1:C3", document.GetTables().Single().Range);

            sheet.SetTableTotalsByName("Orders", totals);
            Assert.Equal("A1:C4", document.GetTables().Single().Range);
        }

        [Fact]
        public void RowsAs_SkipsTheTotalsRowOfATableEndingTheRange() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateOrdersTable(document, 4);
            sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                ["Amount"] = ExcelTableTotalsFunction.Sum
            });

            TotalsOrder[] usedRange = sheet.RowsAs<TotalsOrder>().ToArray();
            TotalsOrder[] tableRange = sheet.RowsAs<TotalsOrder>("A1:C6").ToArray();
            Assert.Equal(4, usedRange.Length);
            Assert.Equal(4, tableRange.Length);
            Assert.Equal(40d, usedRange[3].Amount);
        }
    
        [Fact]
        public void AddTable_RightAlignsHeadersOverNumericColumns() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateOrdersTable(document, 3);

            Assert.True(string.IsNullOrEmpty(sheet.Cell(1, 1).GetStyle().HorizontalAlignment));
            Assert.Equal("right", sheet.Cell(1, 3).GetStyle().HorizontalAlignment);
        }

        [Fact]
        public void AddTable_GivesWideTablesLandscapePrintSettingsAndRepeatedHeaders() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Wide");
            for (int column = 1; column <= 16; column++) {
                sheet.Cell(1, column).SetValue("Column " + column);
                sheet.Cell(2, column).SetValue(column);
            }

            sheet.AddTable("A1:P2", true, "Wide", ExcelTableStyle.TableStyleMedium2);

            ExcelSheetPageSetup pageSetup = sheet.GetPageSetup()!;
            Assert.Equal(OfficePageOrientation.Landscape, pageSetup.Orientation);
            Assert.Equal(1U, pageSetup.FitToWidth);
            ExcelPrintTitles titles = sheet.GetPrintTitles();
            Assert.Equal(1, titles.FirstRow);
            Assert.Equal(1, titles.LastRow);

            // The caller's choice wins and stops further defaults.
            sheet.SetOrientation(OfficePageOrientation.Portrait);
            sheet.AutoFitColumns();
            Assert.Equal(OfficePageOrientation.Portrait, sheet.GetPageSetup()!.Orientation);
        }

        [Fact]
        public void AddTable_LeavesNarrowTablesWithoutPageSetup() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreateOrdersTable(document, 2);
            Assert.Null(sheet.GetPageSetup()?.Orientation);
        }

        [Fact]
        public void SetTableTotals_SecondCallKeepsGeneratedCellsInStepAndUserFormulasIntact() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = CreateOrdersTable(document, 3);
                sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                    ["Amount"] = ExcelTableTotalsFunction.Sum,
                    ["Date"] = ExcelTableTotalsFunction.Maximum
                });
                sheet.Cell(5, 2).SetFormula("COUNT(Orders[Date])");
                Assert.Throws<ArgumentException>(() => sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                    ["Amount"] = ExcelTableTotalsFunction.Custom
                }));
                sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                    ["Customer"] = ExcelTableTotalsFunction.Count,
                    ["Date"] = ExcelTableTotalsFunction.Minimum,
                    ["Amount"] = ExcelTableTotalsFunction.Average
                });
                document.Save();
            }

            stream.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(stream, false);
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            var cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("SUBTOTAL(103,Orders[Customer])", cells["A5"].CellFormula!.Text);
            Assert.Equal("COUNT(Orders[Date])", cells["B5"].CellFormula!.Text);
            Assert.Equal("SUBTOTAL(101,Orders[Amount])", cells["C5"].CellFormula!.Text);
            Assert.Equal("20", cells["C5"].CellValue!.Text);
            var column = worksheetPart.TableDefinitionParts.Single().Table!.TableColumns!
                .Elements<DocumentFormat.OpenXml.Spreadsheet.TableColumn>().First();
            Assert.Null(column.TotalsRowLabel);
        }

        [Fact]
        public void SetTableTotals_NoneClearsTheGeneratedFormula() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = CreateOrdersTable(document, 2);
                sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> { ["Amount"] = ExcelTableTotalsFunction.Sum });
                sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> { ["Amount"] = ExcelTableTotalsFunction.None });
                document.Save();
            }

            stream.Position = 0;
            using var spreadsheet = SpreadsheetDocument.Open(stream, false);
            Assert.DoesNotContain(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "C4" && (cell.CellFormula != null || cell.CellValue != null));
        }
    }
}
