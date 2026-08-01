using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_MutationDiagnostics_ObserveCancellation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Data").CellValue(1, 1, "value");
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();

            Assert.ThrowsAny<OperationCanceledException>(() =>
                document.GetMutationDiagnostics(100, cancellation.Token));
        }

        [Fact]
        public async Task Test_QueryBackedTable_RefreshPreservesTotalsRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "TotalsQuery",
                WorksheetName = sheet.Name,
                TableName = "TotalsResults",
                ColumnNames = new[] { "Region", "Amount" }
            });
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            table.Reference = "A1:B3";
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            table.AutoFilter!.Reference = "A1:B2";
            TableColumn[] columns = table.TableColumns!.Elements<TableColumn>().ToArray();
            columns[0].TotalsRowLabel = "Total";
            columns[1].TotalsRowFormula = new TotalsRowFormula("SUBTOTAL(109,[Amount])");
            sheet.CellValue(2, 1, "old");
            sheet.CellValue(2, 2, 1d);
            sheet.CellValue(3, 1, "Total");
            sheet.CellAt(3, 2).SetFormula("SUBTOTAL(109,[Amount])");
            table.Save();

            var host = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "Region", "Amount" },
                new IReadOnlyList<object?>[] {
                    new object?[] { "East", 10d },
                    new object?[] { "West", 20d }
                }));
            ExcelQueryRefreshResult refreshed = await document.RefreshQueryAsync(
                source.TableName,
                host,
                new ExcelQueryExecutionPolicy { AllowExecution = true });

            Assert.Equal("A1:B4", refreshed.Range);
            Assert.Equal("A1:B4", table.Reference!.Value);
            Assert.Equal("A1:B3", table.AutoFilter!.Reference!.Value);
            Assert.Equal("Total", sheet.CellAt(4, 1).GetValue<string>());
            Assert.Equal("SUBTOTAL(109,[Amount])", sheet.GetFormulaText(4, 2));
            Assert.Equal("Total", columns[0].TotalsRowLabel!.Value);
            Assert.Equal("SUBTOTAL(109,[Amount])", columns[1].TotalsRowFormula!.Text);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernChart_GrowingUpdatesReuseAndReleaseOwnedDataRows() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                ExcelModernChart chart = sheet.AddModernChart(
                    CreateModernChartData(1),
                    row: 2,
                    column: 2,
                    ExcelModernChartType.Funnel);
                int maximumStartRow = chart.DataRange!.StartRow;
                for (int count = 2; count <= 40; count++) {
                    chart.UpdateData(CreateModernChartData(count));
                    maximumStartRow = Math.Max(maximumStartRow, chart.DataRange!.StartRow);
                }
                Assert.True(maximumStartRow < 250, $"Chart data allocation reached row {maximumStartRow}.");
                document.Save(stream);
            }

            stream.Position = 0;
            using (var loaded = ExcelDocument.Load(stream)) {
                ExcelModernChart chart = Assert.Single(loaded["Dashboard"].ModernCharts);
                Assert.True(chart.DataRange!.HasHeaderRow);
                chart.UpdateData(CreateModernChartData(41));
                ExcelSheet dataSheet = loaded[chart.DataRange.SheetName];
                chart.Remove();
                Assert.DoesNotContain(dataSheet.WorksheetPart.Worksheet.Descendants<Cell>(), cell =>
                    cell.CellFormula != null || cell.CellValue != null || cell.InlineString != null);
                Assert.Empty(loaded.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_RangeMutationsAndClearResolveImplicitCellAddressesWithoutMutatingDryRun() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var row = new Row(
                new Cell { CellValue = new CellValue("A"), DataType = CellValues.String },
                new Cell { CellValue = new CellValue("B"), DataType = CellValues.String });
            sheetData.Append(row);

            ExcelStructuralMutationPlan plan = sheet.PlanCopyRange("A1:B1", "C1");
            Assert.Equal(2, Assert.Single(plan.Impacts, impact => impact.Category == "cells").ItemCount);
            Assert.Null(row.RowIndex);
            Assert.All(row.Elements<Cell>(), cell => Assert.Null(cell.CellReference));

            plan.Apply();
            Assert.Equal("A", sheet.CellAt(1, 3).GetValue<string>());
            Assert.Equal("B", sheet.CellAt(1, 4).GetValue<string>());

            var implicitRow = new Row(
                new Cell { CellValue = new CellValue("clear"), DataType = CellValues.String },
                new Cell { CellValue = new CellValue("keep"), DataType = CellValues.String });
            sheetData.Append(implicitRow);
            sheet.ClearRange("A2:A2", ExcelClearOptions.Values);
            Assert.Null(implicitRow.Elements<Cell>().First().CellValue);
            Assert.Equal("keep", implicitRow.Elements<Cell>().Last().CellValue!.Text);

            sheet.InsertCells("A2:A2", ExcelCellShiftDirection.Right);
            Assert.Equal("keep", sheet.CellAt(2, 3).GetValue<string>());
        }

        private static ExcelChartData CreateModernChartData(int count) {
            string[] categories = Enumerable.Range(1, count).Select(index => "Item " + index).ToArray();
            double[] values = Enumerable.Range(1, count).Select(index => (double)index).ToArray();
            return new ExcelChartData(categories, new[] { new ExcelChartSeries("Value", values) });
        }
    }
}
