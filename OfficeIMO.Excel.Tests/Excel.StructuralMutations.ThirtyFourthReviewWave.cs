using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernChart_RejectsInvalidXmlTextBeforeAnyMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            var validData = new ExcelChartData(
                new[] { "A" },
                new[] { new ExcelChartSeries("Value", new[] { 1d }) });
            var invalidData = new ExcelChartData(
                new[] { "Bad\u0001Category" },
                new[] { new ExcelChartSeries("Value", new[] { 2d }) });

            Assert.Throws<ArgumentException>(() => sheet.AddModernChart(
                invalidData,
                1,
                1,
                ExcelModernChartType.Funnel));
            Assert.Throws<ArgumentException>(() => sheet.AddModernChart(
                validData,
                1,
                1,
                ExcelModernChartType.Funnel,
                title: "Bad\u0001Title"));
            Assert.Empty(sheet.ModernCharts);
            Assert.Single(document.Sheets);

            ExcelModernChart chart = sheet.AddModernChart(
                validData,
                1,
                1,
                ExcelModernChartType.Funnel,
                title: "Original");
            ExcelChartDataRange range = chart.DataRange!;
            double originalValue = document[range.SheetName]
                .CellAt(range.CategoryStartRow, range.SeriesStartColumn)
                .GetValue<double>();
            var invalidUpdate = new ExcelChartData(
                new[] { "A" },
                new[] { new ExcelChartSeries("Bad\u0001Series", new[] { 9d }) });

            Assert.Throws<ArgumentException>(() => chart.UpdateData(invalidUpdate));
            Assert.Equal(originalValue, document[range.SheetName]
                .CellAt(range.CategoryStartRow, range.SeriesStartColumn)
                .GetValue<double>());
            Assert.Throws<ArgumentException>(() => chart.SetTitle("Bad\u0001Title"));
            Assert.Equal("Original", chart.Title);
            string originalName = chart.Name;
            Assert.Throws<ArgumentException>(() => chart.Name = "Bad\u0001Name");
            Assert.Equal(originalName, chart.Name);
        }

        [Fact]
        public async Task Test_QueryBackedTable_BudgetsDateTimeOffsetFallbackTextBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelQueryBackedTableInfo source = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "Dates",
                WorksheetName = sheet.Name,
                TableName = "DateResults",
                ColumnNames = new[] { "When" }
            });
            DateTimeOffset fallback = DateTimeOffset.MinValue;
            string fallbackText = fallback.ToString("o", CultureInfo.InvariantCulture);
            var host = new StubQueryHost(new ExcelQueryExecutionResult(
                new[] { "When" },
                new IReadOnlyList<object?>[] { new object?[] { fallback } }));

            await Assert.ThrowsAsync<InvalidOperationException>(() => document.RefreshQueryAsync(
                source.TableName,
                host,
                new ExcelQueryExecutionPolicy {
                    AllowExecution = true,
                    MaximumCharacters = "When".Length + fallbackText.Length - 1L
            }));
            Assert.Equal("When", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Null(sheet.CellAt(2, 1).GetValue<string>());

            await document.RefreshQueryAsync(
                source.TableName,
                host,
                new ExcelQueryExecutionPolicy {
                    AllowExecution = true,
                    MaximumCharacters = "When".Length + fallbackText.Length
                });
            Assert.Equal(fallbackText, sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_FeatureReport_HandlesChartExWithMissingRelationship() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            DrawingsPart drawingsPart = sheet.WorksheetPart.DrawingsPart!;
            drawingsPart.DeletePart(Assert.Single(drawingsPart.ExtendedChartParts));

            ExcelModernChart malformed = Assert.Single(sheet.ModernCharts);
            Assert.Equal(ExcelModernChartType.Unsupported, malformed.ChartType);
            ExcelFeatureReport report = document.InspectFeatures();

            ExcelFeatureFinding charts = Assert.Single(report.FindFeatures("Charts"));
            Assert.Equal(1, charts.Count);
            Assert.Contains(
                report.GetCapabilityDiagnostics(ExcelPreflightCapability.ExportPdfReport),
                diagnostic => diagnostic.Contains("Unsupported ChartEx", StringComparison.OrdinalIgnoreCase));
        }
    }
}
