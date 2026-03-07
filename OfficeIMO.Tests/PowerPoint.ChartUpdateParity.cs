using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointChartUpdateParity {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CanUpdatePieLikeChartData(bool doughnut) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var initialData = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] { new PowerPointChartSeries("Original", new[] { 1d, 2d, 3d }) });
            var updatedData = new PowerPointChartData(
                new[] { "North", "South", "West" },
                new[] { new PowerPointChartSeries("Updated", new[] { 10d, 20d, 30d }) });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddChart(initialData);
                    presentation.Save();
                }

                ConvertFirstChartToPieLike(filePath, doughnut);

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointChart chart = presentation.Slides.Last().Charts.Single();
                    chart.UpdateData(updatedData);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
                    OpenXmlCompositeElement chartElement = doughnut
                        ? plotArea.GetFirstChild<C.DoughnutChart>()!
                        : plotArea.GetFirstChild<C.PieChart>()!;
                    var series = chartElement.Elements<C.PieChartSeries>().Single();

                    string? title = series.GetFirstChild<C.SeriesText>()?
                        .GetFirstChild<C.StringReference>()?
                        .GetFirstChild<C.StringCache>()?
                        .Elements<C.StringPoint>()
                        .Select(point => point.NumericValue?.Text)
                        .SingleOrDefault();
                    Assert.Equal("Updated", title);

                    string[] categories = series.GetFirstChild<C.CategoryAxisData>()?
                        .GetFirstChild<C.StringReference>()?
                        .GetFirstChild<C.StringCache>()?
                        .Elements<C.StringPoint>()
                        .Select(point => point.NumericValue?.Text ?? string.Empty)
                        .ToArray() ?? Array.Empty<string>();
                    Assert.Equal(new[] { "North", "South", "West" }, categories);

                    double[] values = series.GetFirstChild<C.Values>()?
                        .GetFirstChild<C.NumberReference>()?
                        .GetFirstChild<C.NumberingCache>()?
                        .Elements<C.NumericPoint>()
                        .Select(point => double.Parse(point.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                        .ToArray() ?? Array.Empty<double>();
                    Assert.Equal(new[] { 10d, 20d, 30d }, values);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanUpdateLineChartDataWithoutCloningTrendlines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var initialData = new PowerPointChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] { new PowerPointChartSeries("Revenue", new[] { 12d, 18d, 15d }) });
            var updatedData = new PowerPointChartData(
                new[] { "Apr", "May", "Jun" },
                new[] {
                    new PowerPointChartSeries("Revenue", new[] { 20d, 22d, 24d }),
                    new PowerPointChartSeries("Forecast", new[] { 18d, 19d, 21d })
                });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddLineChart(initialData);
                    chart.SetSeriesTrendline(0, C.TrendlineValues.Linear, lineColor: "ED7D31", lineWidthPoints: 1.25)
                        .UpdateData(updatedData);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.LineChartSeries[] series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.LineChart>()!
                        .Elements<C.LineChartSeries>()
                        .ToArray();

                    Assert.Equal(2, series.Length);
                    Assert.NotNull(series[0].GetFirstChild<C.Trendline>());
                    Assert.Null(series[1].GetFirstChild<C.Trendline>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void ConvertFirstChartToPieLike(string filePath, bool doughnut) {
            using PresentationDocument document = PresentationDocument.Open(filePath, true);
            ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
            C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            C.BarChart barChart = plotArea.GetFirstChild<C.BarChart>()!;

            OpenXmlCompositeElement replacementChart = doughnut
                ? new C.DoughnutChart(new C.VaryColors { Val = true })
                : new C.PieChart(new C.VaryColors { Val = true });

            foreach (C.BarChartSeries barSeries in barChart.Elements<C.BarChartSeries>()) {
                replacementChart.Append(CreatePieChartSeries(barSeries));
            }

            if (barChart.GetFirstChild<C.DataLabels>() is C.DataLabels labels) {
                replacementChart.Append(labels.CloneNode(true));
            }

            if (doughnut) {
                replacementChart.Append(new C.HoleSize { Val = 50 });
            }

            barChart.Remove();
            plotArea.RemoveAllChildren<C.CategoryAxis>();
            plotArea.RemoveAllChildren<C.ValueAxis>();
            plotArea.Append(replacementChart);
            chartPart.ChartSpace.Save();
        }

        private static C.PieChartSeries CreatePieChartSeries(C.BarChartSeries barSeries) {
            C.PieChartSeries pieSeries = new();

            if (barSeries.GetFirstChild<C.Index>() is C.Index index) {
                pieSeries.Append((C.Index)index.CloneNode(true));
            }
            if (barSeries.GetFirstChild<C.Order>() is C.Order order) {
                pieSeries.Append((C.Order)order.CloneNode(true));
            }
            if (barSeries.GetFirstChild<C.SeriesText>() is C.SeriesText seriesText) {
                pieSeries.Append((C.SeriesText)seriesText.CloneNode(true));
            }
            if (barSeries.GetFirstChild<C.CategoryAxisData>() is C.CategoryAxisData categoryAxisData) {
                pieSeries.Append((C.CategoryAxisData)categoryAxisData.CloneNode(true));
            }
            if (barSeries.GetFirstChild<C.Values>() is C.Values values) {
                pieSeries.Append((C.Values)values.CloneNode(true));
            }

            return pieSeries;
        }
    }
}
