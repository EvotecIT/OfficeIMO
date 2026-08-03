using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_DiscardedCells_ReclaimExclusiveInCellImagesWhileMovesPreserveThem() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Ranges");
                sheet.CellValue(1, 1, "Replacement");
                sheet.SetInCellImage(1, 2, TinyPng, altText: "Discarded destination");

                sheet.Range("A1").CopyTo("B1");

                Assert.Empty(sheet.GetInCellImages());
                Assert.Empty(document.WorkbookPartRoot.RdRichValueParts);

                sheet.SetInCellImage(1, 3, TinyPng, altText: "Moved source");
                sheet.MoveRange("C1", "D1");

                ExcelInCellImage moved = Assert.Single(sheet.GetInCellImages());
                Assert.Equal("D1", moved.CellReference);
                Assert.Equal(TinyPng, moved.Bytes);
                AssertRichImageGraphCounts(document, expected: 1);

                sheet.SetInCellImage(5, 5, TinyPng, altText: "Deleted row");
                sheet.DeleteRows(5);
                sheet.SetInCellImage(6, 6, TinyPng, altText: "Deleted column");
                sheet.DeleteColumns(6);
                sheet.SetInCellImage(7, 7, TinyPng, altText: "Deleted cell");
                sheet.DeleteCells("G7", ExcelCellShiftDirection.Left);

                ExcelInCellImage survivor = Assert.Single(sheet.GetInCellImages());
                Assert.Equal("D1", survivor.CellReference);
                AssertRichImageGraphCounts(document, expected: 1);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Query");
                ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                    ConnectionName = "RefreshImages",
                    WorksheetName = sheet.Name,
                    TableName = "RefreshResults",
                    ColumnNames = new[] { "Value" }
                });
                sheet.SetInCellImage(1, 1, TinyPng, altText: "Replaced header");
                var host = new StubQueryHost(new ExcelQueryExecutionResult(
                    new[] { "Value" },
                    new IReadOnlyList<object?>[] { new object?[] { "Updated" } }));

                await document.RefreshQueryAsync(
                    source.TableName,
                    host,
                    new ExcelQueryExecutionPolicy { AllowExecution = true });

                Assert.Empty(sheet.GetInCellImages());
                Assert.Empty(document.WorkbookPartRoot.RdRichValueParts);
                Assert.Equal("Value", sheet.CellAt(1, 1).GetValue<string>());
                Assert.Equal("Updated", sheet.CellAt(2, 1).GetValue<string>());
            }
        }

        [Fact]
        public void Test_OwnedChartDataWorksheet_CannotBeRemovedWhileChartsReferenceIt() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet dashboard = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = dashboard.AddModernChart(
                new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            string dataSheetName = chart.DataRange!.SheetName;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                document.RemoveWorksheet(dataSheetName));

            Assert.Contains("referenced by OfficeIMO-authored charts", exception.Message, StringComparison.Ordinal);
            chart.UpdateData(new ExcelChartData(
                new[] { "A", "B" },
                new[] { new ExcelChartSeries("Value", new[] { 3d, 4d }) }));
            Assert.Contains(document.Sheets, sheet => sheet.Name == dataSheetName);

            chart.Remove();
            document.RemoveWorksheet(dataSheetName);
            Assert.DoesNotContain(document.Sheets, sheet => sheet.Name == dataSheetName);

            ExcelModernChart replacement = dashboard.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 5d }) }),
                1,
                1,
                ExcelModernChartType.Treemap);
            Assert.True(document[replacement.DataRange!.SheetName].Hidden);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_TableSchema_RejectsOversizedHeadersBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Original");
            sheet.CellValue(2, 1, "Keep");
            sheet.AddTable("A1:A2", true, "DataTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            string originalXml = table.OuterXml;

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                sheet.SetTableSchema("DataTable", new[] { new string('X', 32_768) }));

            Assert.Equal("columnNames", exception.ParamName);
            Assert.Equal(originalXml, table.OuterXml);
            Assert.Equal("Original", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("Keep", sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_FeatureReport_PreservesUnresolvableQueryTableParts() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "BrokenBinding",
                WorksheetName = sheet.Name,
                TableName = "BrokenResults",
                ColumnNames = new[] { "Value" }
            });
            QueryTablePart queryPart = Assert.Single(
                Assert.Single(sheet.WorksheetPart.TableDefinitionParts).QueryTableParts);
            queryPart.QueryTable!.ConnectionId = null;
            queryPart.QueryTable.Save();

            ExcelFeatureReport report = document.InspectFeatures();

            Assert.Empty(report.FindFeatures("Query-backed tables"));
            ExcelFeatureFinding preserved = Assert.Single(report.FindFeatures("Connections and query tables"));
            Assert.Equal(2, preserved.Count);
            Assert.Contains(preserved.Details, detail => detail.Contains("queryTable", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(preserved.Details, detail => detail.Contains("connection", StringComparison.OrdinalIgnoreCase));
        }
    }
}
