using System.Linq;
using Xunit;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_PiePalette_SemanticOutcomes() {
            var filePath = Path.Combine(_directoryWithFiles, "Chart.Pie.SemanticPalette.docx");

            using (var doc = WordDocument.Create(filePath)) {
                var pie = doc.AddChart("Rules outcome", false, 600, 320);
                pie.AddPie("Passed", 42);
                pie.AddPie("Failed", 30);
                pie.AddPie("Skipped", 5);
                pie.AddLegend(LegendPositionValues.Right);

                pie.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false)
                   .SetWidthToPageContent(1.0, 320);

                doc.Save(false);
            }

            using (var doc = WordDocument.Load(filePath)) {
                var part = doc._wordprocessingDocument!.MainDocumentPart!.ChartParts.First();
                var chart = part.ChartSpace.GetFirstChild<Chart>()!;
                var pie = chart.PlotArea!.GetFirstChild<PieChart>()!;
                var series = pie.GetFirstChild<PieChartSeries>()!;
                var dpts = series.Elements<DataPoint>().OrderBy(d => d.Index!.Val!.Value).ToList();

                // Expected semantic colors (hex)
                var passed = Color.ParseHex("#2fb344").ToHexColor();
                var failed = Color.ParseHex("#f76707").ToHexColor();
                var skipped = Color.ParseHex("#868e96").ToHexColor();

                string ColorOf(DataPoint dpt)
                    => dpt.GetFirstChild<ChartShapeProperties>()!
                           .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                           .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()!.Val!;

                Assert.Equal(passed, ColorOf(dpts[0]));
                Assert.Equal(failed, ColorOf(dpts[1]));
                Assert.Equal(skipped, ColorOf(dpts[2]));

                // Validation should pass
                var validation = doc.ValidateDocument();
                var chartErrors = validation.Where(v => v.Description.Contains("chart")).ToList();
                Assert.True(chartErrors.Count == 0, Word.FormatValidationErrors(chartErrors));
            }
        }

        [Fact]
        public void Test_BarPalette_ColorBlindSafe() {
            var filePath = Path.Combine(_directoryWithFiles, "Chart.Bar.ColorBlindSafe.docx");

            using (var doc = WordDocument.Create(filePath)) {
                var categories = new[] { "Q1", "Q2", "Q3", "Q4" }.ToList();
                var bar = doc.AddChart("Quarterly", false, 600, 320);
                bar.AddCategories(categories);
                bar.AddBar("EMEA", new[] { 10, 12, 14, 18 }, Color.Black);
                bar.AddBar("APAC", new[] { 9, 11, 15, 20 }, Color.Black);
                bar.AddBar("AMER", new[] { 8, 10, 16, 19 }, Color.Black);

                bar.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe)
                   .SetWidthToPageContent(1.0, 320);

                doc.Save(false);
            }

            using (var doc = WordDocument.Load(filePath)) {
                var part = doc._wordprocessingDocument!.MainDocumentPart!.ChartParts.First();
                var chart = part.ChartSpace.GetFirstChild<Chart>()!;
                var bar = chart.PlotArea!.GetFirstChild<BarChart>()!;
                var series = bar.Elements<BarChartSeries>().OrderBy(s => s.Index!.Val!.Value).ToList();

                string ColorOf(BarChartSeries s)
                    => s.GetFirstChild<ChartShapeProperties>()!
                         .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                         .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()!.Val!;

                // First three Okabe–Ito colors, normalized to hex without '#'
                var expected = new[] {
                    Color.ParseHex("#0072B2").ToHexColor(),
                    Color.ParseHex("#E69F00").ToHexColor(),
                    Color.ParseHex("#009E73").ToHexColor()
                };
                Assert.Equal(expected[0], ColorOf(series[0]));
                Assert.Equal(expected[1], ColorOf(series[1]));
                Assert.Equal(expected[2], ColorOf(series[2]));

                var validation = doc.ValidateDocument();
                var chartErrors = validation.Where(v => v.Description.Contains("chart")).ToList();
                Assert.True(chartErrors.Count == 0, Word.FormatValidationErrors(chartErrors));
            }
        }

        [Fact]
        public void Test_FitToPageContentWidth_SetsInlineExtent() {
            var filePath = Path.Combine(_directoryWithFiles, "Chart.FitWidth.docx");

            using (var doc = WordDocument.Create(filePath)) {
                var ch = doc.AddChart("Full width", false, 400, 240);
                ch.SetWidthToPageContent(1.0, 240);
                doc.Save(false);
            }

            using (var doc = WordDocument.Load(filePath)) {
                // Compute expected EMUs from page size (Letter) and Normal margins
                var sect = doc.Sections.First();
                double widthTwips = sect.PageSettings.Width!.Value!; // twips
                double leftTwips = sect.Margins.Left!.Value!;
                double rightTwips = sect.Margins.Right!.Value!;
                double contentInches = (widthTwips - leftTwips - rightTwips) / 1440.0;
                long expectedCx = (long)System.Math.Round(contentInches * 914400); // EMUs
                long expectedCy = (long)System.Math.Round((240.0 / 96.0) * 914400); // 240px -> inches -> EMUs

                var inline = doc._wordprocessingDocument!.MainDocumentPart!.Document!
                    .Body!.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().First();
                var extent = inline.Extent;
                Assert.NotNull(extent);
                var cx = extent!.Cx!.Value!;
                var cy = extent!.Cy!.Value!;

                Assert.Equal(expectedCx, cx);
                Assert.Equal(expectedCy, cy);
            }
        }

        [Fact]
        public void Test_SetSeriesColor_OverridesPalette() {
            var filePath = Path.Combine(_directoryWithFiles, "Chart.SeriesColorOverride.docx");

            using (var doc = WordDocument.Create(filePath)) {
                var categories = new[] { "A", "B", "C" }.ToList();
                var line = doc.AddChart("KPIs", false, 600, 320);
                line.AddChartAxisX(categories);
                line.AddLine("Throughput", new[] { 1, 2, 3 }.ToList(), Color.Black);
                line.AddLine("Latency", new[] { 3, 2, 1 }.ToList(), Color.Black);
                line.ApplyPalette(WordChart.WordChartPalette.MonochromeGray)
                    .SetSeriesColor(1, Color.ParseHex("#d63939"));
                doc.Save(false);
            }

            using (var doc = WordDocument.Load(filePath)) {
                var part = doc._wordprocessingDocument!.MainDocumentPart!.ChartParts.First();
                var c = part.ChartSpace.GetFirstChild<Chart>()!;
                var line = c.PlotArea!.GetFirstChild<LineChart>()!;
                var series = line.Elements<LineChartSeries>().OrderBy(s => s.Index!.Val!.Value).ToList();
                var s1 = series[1];
                var hex = s1.GetFirstChild<ChartShapeProperties>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()!.Val!;
                Assert.Equal(Color.ParseHex("#d63939").ToHexColor(), hex);

                var validation = doc.ValidateDocument();
                var chartErrors = validation.Where(v => v.Description.Contains("chart")).ToList();
                Assert.True(chartErrors.Count == 0, Word.FormatValidationErrors(chartErrors));
            }
        }
    }
}
