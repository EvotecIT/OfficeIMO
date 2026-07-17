using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

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
                    new[] { 100.25d, 40.5d, -30.1d },
                    row: 20,
                    column: 10);
                Assert.Equal(ExcelChartType.ColumnStacked, waterfall.ChartType);
                Assert.True(waterfall.TryGetData(out ExcelChartData waterfallData));
                Assert.Equal(new[] { "Start", "Growth", "Cost", "Total" }, waterfallData.Categories);
                Assert.Equal(new[] { 0d, 100.25d, 110.65d, 0d }, waterfallData.Series[0].Values);
                Assert.Equal(new[] { 100.25d, 40.5d, 0d, 0d }, waterfallData.Series[1].Values);
                Assert.Equal(new[] { 0d, 0d, 30.1d, 0d }, waterfallData.Series[2].Values);
                Assert.Equal(new[] { 0d, 0d, 0d, 110.65d }, waterfallData.Series[3].Values);

                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, false)) {
                C.BarChart waterfallChart = spreadsheet.WorkbookPart!.WorksheetParts
                    .Where(part => part.DrawingsPart != null)
                    .SelectMany(part => part.DrawingsPart!.ChartParts)
                    .SelectMany(part => part.ChartSpace.Descendants<C.BarChart>())
                    .Single(chart => chart.Elements<C.BarChartSeries>().Count() == 4);
                C.BarChartSeries[] waterfallSeries = waterfallChart.Elements<C.BarChartSeries>().ToArray();
                Assert.Null(waterfallSeries[0].GetFirstChild<C.DataLabels>());
                Assert.Equal(new uint[] { 0, 1 }, waterfallSeries[1].GetFirstChild<C.DataLabels>()!
                    .Elements<C.DataLabel>().Select(label => label.Index!.Val!.Value));
                Assert.All(waterfallSeries[1].GetFirstChild<C.DataLabels>()!.Elements<C.DataLabel>(), label =>
                    Assert.Equal("+#,##0.###############", label.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value));
                Assert.Equal(new uint[] { 2 }, waterfallSeries[2].GetFirstChild<C.DataLabels>()!
                    .Elements<C.DataLabel>().Select(label => label.Index!.Val!.Value));
                Assert.Equal("-#,##0.###############", waterfallSeries[2].GetFirstChild<C.DataLabels>()!
                    .GetFirstChild<C.DataLabel>()!.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                Assert.Equal(new uint[] { 3 }, waterfallSeries[3].GetFirstChild<C.DataLabels>()!
                    .Elements<C.DataLabel>().Select(label => label.Index!.Val!.Value));
                Assert.Equal("#,##0.###############", waterfallSeries[3].GetFirstChild<C.DataLabels>()!
                    .GetFirstChild<C.DataLabel>()!.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
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
            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { -double.MaxValue, double.MaxValue }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { 0d, double.Epsilon }, 1, 1, binCount: 2));
            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { 1e16d, 10000000000000002d }, 1, 1, binWidth: 1d));
            Assert.Throws<ArgumentException>(() => sheet.AddHistogramChart(new[] { 1d, 2d }, 1, 1, binCount: 2, binWidth: 1));
            Assert.Throws<ArgumentException>(() => sheet.AddParetoChart(new[] { "A", "B" }, new[] { 0d, 0d }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddParetoChart(new[] { "A", "B" }, new[] { double.MaxValue, double.MaxValue }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddFunnelChart(new[] { "A" }, new[] { -1d }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddWaterfallChart(new[] { "A" }, new[] { -1d }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddWaterfallChart(new[] { "A", "B" }, new[] { double.MaxValue, double.MaxValue }, 1, 1));
            Assert.Throws<ArgumentException>(() => sheet.AddWaterfallChart(
                new[] { "A", "B" },
                new[] { 1e16d, -10000000000000002d },
                1,
                1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddHistogramChart(new[] { 1d, 2d }, 1, 1, widthPixels: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddParetoChart(new[] { "A" }, new[] { 1d }, 1, 1, heightPixels: 0));
            Assert.Single(document.Sheets);

            ExcelChart wideBinHistogram = sheet.AddHistogramChart(
                new[] { 0d, double.Epsilon },
                row: 1,
                column: 1,
                binWidth: double.MaxValue);
            Assert.True(wideBinHistogram.TryGetData(out ExcelChartData wideBinData));
            Assert.Single(wideBinData.Categories);
            Assert.Equal(2d, Assert.Single(wideBinData.Series).Values.Single());

            ExcelChart tinyRangeHistogram = sheet.AddHistogramChart(
                new[] { 0d, 1e-11d },
                row: 1,
                column: 10,
                binCount: 2);
            Assert.True(tinyRangeHistogram.TryGetData(out ExcelChartData tinyRangeData));
            Assert.Equal(2, tinyRangeData.Categories.Count);
            Assert.Equal(2, tinyRangeData.Categories.Distinct(StringComparer.Ordinal).Count());
            Assert.DoesNotContain("0 – 0", tinyRangeData.Categories);

            ExcelChart roundingSafeWaterfall = sheet.AddWaterfallChart(
                new[] { "Start", "First", "Second" },
                new[] { 0.3d, -0.1d, -0.2d },
                row: 10,
                column: 1);
            Assert.True(roundingSafeWaterfall.TryGetData(out ExcelChartData roundingSafeData));
            Assert.Equal(0d, roundingSafeData.Series[3].Values.Last());

            ExcelChart noTotalWaterfall = sheet.AddWaterfallChart(
                new[] { "Start", "Cost" },
                new[] { 2d, -1d },
                row: 10,
                column: 10,
                includeTotal: false);
            Assert.True(noTotalWaterfall.TryGetData(out ExcelChartData noTotalData));
            Assert.Equal(3, noTotalData.Series.Count);
        }

        [Fact]
        public void Test_SetSeriesNoFill_NormalizesLoadedOutlineFillChoice() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernChartRecipes.NoFill.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                var data = new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }, seriesColorArgb: "4F46E5") });
                sheet.AddChart(data, 1, 1, 500, 300, ExcelChartType.ColumnClustered, "Loaded outline")
                    .SetSeriesNoFill(0, noLine: false);
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, true)) {
                DocumentFormat.OpenXml.Packaging.ChartPart chartPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single();
                C.BarChartSeries series = chartPart.ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .Single();
                C.ChartShapeProperties properties = series.GetFirstChild<C.ChartShapeProperties>()!;
                properties.Append(new A.BlipFill());
                properties.Append(new A.GroupFill());
                properties.GetFirstChild<A.Outline>()?.Remove();
                properties.Append(new A.Outline(
                    new A.GradientFill(),
                    new A.PatternFill(),
                    new A.PresetDash { Val = A.PresetLineDashValues.Dash }));
                chartPart.ChartSpace.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Single(document["Dashboard"].Charts).SetSeriesNoFill(0);
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, false)) {
                A.Outline outline = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single()
                    .ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .Single()
                    .GetFirstChild<C.ChartShapeProperties>()!
                    .GetFirstChild<A.Outline>()!;
                Assert.Single(outline.Elements<A.NoFill>());
                Assert.Empty(outline.Elements<A.SolidFill>());
                Assert.Empty(outline.Elements<A.GradientFill>());
                Assert.Empty(outline.Elements<A.PatternFill>());
                Assert.IsType<A.NoFill>(outline.FirstChild);
                C.ChartShapeProperties properties = Assert.IsType<C.ChartShapeProperties>(outline.Parent);
                Assert.Empty(properties.Elements<A.BlipFill>());
                Assert.Empty(properties.Elements<A.GroupFill>());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_SetSeriesNoFill_InsertsMissingOutlineBeforeEffects() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernChartRecipes.NoFillEffects.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                var data = new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }, seriesColorArgb: "4F46E5") });
                sheet.AddChart(data, 1, 1, 500, 300, ExcelChartType.ColumnClustered, "Effect ordering")
                    .SetSeriesNoFill(0, noLine: false);
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, true)) {
                C.ChartShapeProperties properties = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single()
                    .ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .Single()
                    .GetFirstChild<C.ChartShapeProperties>()!;
                properties.GetFirstChild<A.Outline>()?.Remove();
                properties.Append(new A.EffectList());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Single(document["Dashboard"].Charts).SetSeriesNoFill(0);
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, false)) {
                C.ChartShapeProperties properties = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single()
                    .ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .Single()
                    .GetFirstChild<C.ChartShapeProperties>()!;
                A.Outline outline = Assert.IsType<A.Outline>(properties.GetFirstChild<A.Outline>());
                A.EffectList effects = Assert.IsType<A.EffectList>(properties.GetFirstChild<A.EffectList>());
                Assert.True(properties.ChildElements.ToList().IndexOf(outline) < properties.ChildElements.ToList().IndexOf(effects));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_SetSeriesLineColor_NormalizesLineFillAndInsertsOutlineBeforeEffects() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernChartRecipes.LineFill.xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                var data = new ExcelChartData(
                    new[] { "A", "B" },
                    new[] {
                        new ExcelChartSeries("First", new[] { 1d, 2d }, seriesColorArgb: "4F46E5"),
                        new ExcelChartSeries("Second", new[] { 2d, 3d }, seriesColorArgb: "16A34A")
                    });
                ExcelChart chart = sheet.AddChart(data, 1, 1, 500, 300, ExcelChartType.ColumnClustered, "Line normalization");
                chart.SetSeriesNoFill(0).SetSeriesLineColor(0, "DC2626", 1.5d);
                chart.SetSeriesLineColor(1, "16A34A");
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, true)) {
                C.BarChartSeries secondSeries = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single()
                    .ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .ElementAt(1);
                C.ChartShapeProperties properties = secondSeries.GetFirstChild<C.ChartShapeProperties>()!;
                properties.GetFirstChild<A.Outline>()?.Remove();
                properties.Append(new A.EffectList());
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.Single(document["Dashboard"].Charts).SetSeriesLineColor(1, "2563EB");
                document.Save();
            }

            using (DocumentFormat.OpenXml.Packaging.SpreadsheetDocument spreadsheet =
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, false)) {
                C.BarChartSeries[] series = spreadsheet.WorkbookPart!.WorksheetParts
                    .Single(part => part.DrawingsPart != null)
                    .DrawingsPart!
                    .ChartParts
                    .Single()
                    .ChartSpace
                    .Descendants<C.BarChartSeries>()
                    .ToArray();
                A.Outline firstOutline = Assert.IsType<A.Outline>(series[0].GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>());
                Assert.Single(firstOutline.Elements<A.SolidFill>());
                Assert.Empty(firstOutline.Elements<A.NoFill>());
                Assert.Empty(firstOutline.Elements<A.GradientFill>());
                Assert.Empty(firstOutline.Elements<A.PatternFill>());

                C.ChartShapeProperties secondProperties = series[1].GetFirstChild<C.ChartShapeProperties>()!;
                A.Outline secondOutline = Assert.IsType<A.Outline>(secondProperties.GetFirstChild<A.Outline>());
                A.EffectList effects = Assert.IsType<A.EffectList>(secondProperties.GetFirstChild<A.EffectList>());
                Assert.True(secondProperties.ChildElements.ToList().IndexOf(secondOutline) < secondProperties.ChildElements.ToList().IndexOf(effects));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
