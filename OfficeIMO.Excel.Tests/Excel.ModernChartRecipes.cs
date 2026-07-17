using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernChartRecipes_CreateInspectableCompatibleCharts() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernChartRecipes.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");

                ExcelChart histogram = sheet.AddHistogramChart(
                    new[] { 1d, 2d, 2d, 3d, 4d, 5d },
                    row: 1,
                    column: 1,
                    binCount: 2);
                Assert.Equal(ExcelChartType.ColumnClustered, histogram.ChartType);
                Assert.True(histogram.TryGetData(out ExcelChartData histogramData));
                Assert.Equal(new[] { "1 – 3", "3 – 5" }, histogramData.Categories);
                Assert.Equal(new[] { 3d, 3d }, Assert.Single(histogramData.Series).Values);

                ExcelChart pareto = sheet.AddParetoChart(
                    new[] { "Minor", "Major", "Medium" },
                    new[] { 1d, 6d, 3d },
                    row: 1,
                    column: 10);
                Assert.True(pareto.TryGetData(out ExcelChartData paretoData));
                Assert.Equal(new[] { "Major", "Medium", "Minor" }, paretoData.Categories);
                Assert.Equal(new[] { 0.6d, 0.9d, 1d }, paretoData.Series[1].Values);
                Assert.Equal(ExcelChartType.Line, paretoData.Series[1].ChartType);
                Assert.Equal(ExcelChartAxisGroup.Secondary, paretoData.Series[1].AxisGroup);

                ExcelChart funnel = sheet.AddFunnelChart(
                    new[] { "Leads", "Qualified", "Won" },
                    new[] { 100d, 60d, 20d },
                    row: 20,
                    column: 1);
                Assert.Equal(ExcelChartType.BarStacked, funnel.ChartType);
                Assert.True(funnel.TryGetData(out ExcelChartData funnelData));
                Assert.Equal(new[] { 0d, 20d, 40d }, funnelData.Series[0].Values);
                Assert.Equal(new[] { 100d, 60d, 20d }, funnelData.Series[1].Values);

                ExcelChart waterfall = sheet.AddWaterfallChart(
                    new[] { "Start", "Growth", "Cost" },
                    new[] { 100d, 40d, -30d },
                    row: 20,
                    column: 10);
                Assert.Equal(ExcelChartType.ColumnStacked, waterfall.ChartType);
                Assert.True(waterfall.TryGetData(out ExcelChartData waterfallData));
                Assert.Equal(new[] { "Start", "Growth", "Cost", "Total" }, waterfallData.Categories);
                Assert.Equal(new[] { 0d, 100d, 110d, 0d }, waterfallData.Series[0].Values);
                Assert.Equal(new[] { 100d, 40d, 0d, 0d }, waterfallData.Series[1].Values);
                Assert.Equal(new[] { 0d, 0d, 30d, 0d }, waterfallData.Series[2].Values);
                Assert.Equal(new[] { 0d, 0d, 0d, 110d }, waterfallData.Series[3].Values);

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Equal(4, document["Dashboard"].Charts.Count());
                var validationErrors = document.ValidateOpenXml();
                Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine, validationErrors));
            }
        }

        [Fact]
        public void Test_ModernChartRecipes_RejectInvalidStatisticalInputs() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernChartRecipes.Invalid.xlsx");

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Dashboard");

            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { 1d, double.NaN }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { 1d, 2d }, 1, 1, binCount: 2, binWidth: 1));
            Assert.Throws<ArgumentException>(() => sheet.AddParetoChart(new[] { "A", "B" }, new[] { 0d, 0d }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddFunnelChart(new[] { "A" }, new[] { -1d }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddWaterfallChart(new[] { "A" }, new[] { -1d }, 1, 1));
        }
    }
}
