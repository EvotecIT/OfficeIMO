using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointChartDataBinding {
        private sealed class MetricRow {
            public string Label { get; set; } = string.Empty;
            public double Current { get; set; }
            public double Target { get; set; }
        }

        [Fact]
        public void CanBuildChartFromObjectData() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            var rows = new[] {
                new MetricRow { Label = "Q1", Current = 10, Target = 12 },
                new MetricRow { Label = "Q2", Current = 12, Target = 11 },
                new MetricRow { Label = "Q3", Current = 9, Target = 13 }
            };

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddChart(
                    rows,
                    r => r.Label,
                    new PowerPointChartSeriesDefinition<MetricRow>("Current", r => r.Current),
                    new PowerPointChartSeriesDefinition<MetricRow>("Target", r => r.Target));
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                var chart = chartPart.ChartSpace.GetFirstChild<Chart>()!;
                var barChart = chart.PlotArea!.GetFirstChild<BarChart>()!;
                var series = barChart.Elements<BarChartSeries>().ToList();
                Assert.Equal(2, series.Count);

                var categories = series[0]
                    .GetFirstChild<CategoryAxisData>()?
                    .GetFirstChild<StringReference>()?
                    .GetFirstChild<StringCache>()?
                    .Elements<StringPoint>()
                    .Select(p => p.NumericValue?.Text ?? string.Empty)
                    .ToList();
                Assert.Equal(new[] { "Q1", "Q2", "Q3" }, categories);

                var currentValues = series[0]
                    .GetFirstChild<Values>()?
                    .GetFirstChild<NumberReference>()?
                    .GetFirstChild<NumberingCache>()?
                    .Elements<NumericPoint>()
                    .Select(p => double.Parse(p.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                    .ToList();
                Assert.Equal(new[] { 10d, 12d, 9d }, currentValues);

                var targetValues = series[1]
                    .GetFirstChild<Values>()?
                    .GetFirstChild<NumberReference>()?
                    .GetFirstChild<NumberingCache>()?
                    .Elements<NumericPoint>()
                    .Select(p => double.Parse(p.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                    .ToList();
                Assert.Equal(new[] { 12d, 11d, 13d }, targetValues);

                var externalData = chartPart.ChartSpace.Elements<ExternalData>().FirstOrDefault();
                Assert.NotNull(externalData);
                string? relId = externalData!.Id?.Value;
                var embedded = chartPart.GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();
                Assert.NotNull(embedded);
                Assert.Equal(relId, chartPart.GetIdOfPart(embedded!));
            }

            File.Delete(filePath);
        }
    }
}
