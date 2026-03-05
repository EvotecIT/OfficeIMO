using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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

        private static void ConvertFirstChartToPieLike(string filePath, bool doughnut) {
            using PresentationDocument document = PresentationDocument.Open(filePath, true);
            ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
            C.PlotArea plotArea = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            C.BarChart barChart = plotArea.GetFirstChild<C.BarChart>()!;

            OpenXmlCompositeElement replacementChart = doughnut
                ? new C.DoughnutChart(new C.VaryColors { Val = true }, new C.HoleSize { Val = 50 })
                : new C.PieChart(new C.VaryColors { Val = true });

            foreach (C.BarChartSeries barSeries in barChart.Elements<C.BarChartSeries>()) {
                replacementChart.Append(CreatePieChartSeries(barSeries));
            }

            if (barChart.GetFirstChild<C.DataLabels>() is C.DataLabels labels) {
                replacementChart.Append(labels.CloneNode(true));
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
