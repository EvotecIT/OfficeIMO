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
            sheet.CellValue(1, 1, "Customer");
            sheet.CellValue(1, 2, "Date");
            sheet.CellValue(1, 3, "Amount");
            for (int index = 0; index < orders; index++) {
                sheet.CellValue(index + 2, 1, "Customer " + (index + 1));
                sheet.CellValue(index + 2, 2, new DateTime(2026, 1, 1).AddDays(index));
                sheet.CellValue(index + 2, 3, 10d * (index + 1));
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
            sheet.CellValue(4, 2, "Notes below the table");

            Assert.Throws<InvalidOperationException>(() => sheet.SetTableTotalsByName("Orders", new Dictionary<string, ExcelTableTotalsFunction> {
                ["Amount"] = ExcelTableTotalsFunction.Sum
            }));
        }

        [Fact]
        public void SetTableTotals_ReusesAReservedBlankLastRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Orders");
            sheet.CellValue(1, 1, "Customer");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "Contoso");
            sheet.CellValue(2, 2, 5d);
            sheet.CellValue(3, 1, "Fabrikam");
            sheet.CellValue(3, 2, 7d);
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

            Assert.True(string.IsNullOrEmpty(sheet.GetCellStyle(1, 1).HorizontalAlignment));
            Assert.Equal("right", sheet.GetCellStyle(1, 3).HorizontalAlignment);
        }

        [Fact]
        public void AddTable_GivesWideTablesLandscapePrintSettingsAndRepeatedHeaders() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Wide");
            for (int column = 1; column <= 16; column++) {
                sheet.CellValue(1, column, "Column " + column);
                sheet.CellValue(2, column, column);
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
    }
}
