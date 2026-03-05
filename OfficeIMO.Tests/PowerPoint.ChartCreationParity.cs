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
    public class PowerPointChartCreationParity {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void CanCreatePieLikeCharts(bool doughnut) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var data = new PowerPointChartData(
                new[] { "North", "South", "West" },
                new[] { new PowerPointChartSeries("Revenue", new[] { 10d, 20d, 30d }) });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = doughnut
                        ? slide.AddDoughnutChart(data)
                        : slide.AddPieChart(data);

                    chart.SetTitle(doughnut ? "Revenue Doughnut" : "Revenue Pie")
                        .SetDataLabels(showValue: true, showPercent: true)
                        .SetSeriesFillColor(0, "4472C4");

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    Assert.Null(plotArea.GetFirstChild<C.BarChart>());
                    Assert.Empty(plotArea.Elements<C.CategoryAxis>());
                    Assert.Empty(plotArea.Elements<C.ValueAxis>());

                    var pieChart = plotArea.GetFirstChild<C.PieChart>();
                    var doughnutChart = plotArea.GetFirstChild<C.DoughnutChart>();
                    Assert.Equal(doughnut, doughnutChart != null);
                    Assert.Equal(!doughnut, pieChart != null);

                    var chartElement = (OpenXmlCompositeElement?)doughnutChart ?? pieChart!;
                    var series = chartElement.Elements<C.PieChartSeries>().Single();

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

                    bool? showPercent = chartElement.GetFirstChild<C.DataLabels>()?
                        .GetFirstChild<C.ShowPercent>()?
                        .Val?.Value;
                    Assert.True(showPercent);

                    string? fillColor = series.GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("4472C4", fillColor);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanCreateLineCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var data = new PowerPointChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] { new PowerPointChartSeries("Revenue", new[] { 12d, 18d, 15d }) });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddLineChart(data);

                    chart.SetTitle("Monthly Revenue")
                        .SetDataLabels(showValue: true)
                        .SetSeriesLineColor(0, "ED7D31", 2.5)
                        .SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 9, fillColor: "ED7D31", lineColor: "A64D13");

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    Assert.Null(plotArea.GetFirstChild<C.BarChart>());
                    Assert.Null(plotArea.GetFirstChild<C.PieChart>());
                    Assert.Null(plotArea.GetFirstChild<C.DoughnutChart>());
                    Assert.Single(plotArea.Elements<C.CategoryAxis>());
                    Assert.Single(plotArea.Elements<C.ValueAxis>());

                    C.LineChart lineChart = plotArea.GetFirstChild<C.LineChart>()!;
                    C.LineChartSeries series = lineChart.Elements<C.LineChartSeries>().Single();

                    string[] categories = series.GetFirstChild<C.CategoryAxisData>()?
                        .GetFirstChild<C.StringReference>()?
                        .GetFirstChild<C.StringCache>()?
                        .Elements<C.StringPoint>()
                        .Select(point => point.NumericValue?.Text ?? string.Empty)
                        .ToArray() ?? Array.Empty<string>();
                    Assert.Equal(new[] { "Jan", "Feb", "Mar" }, categories);

                    double[] values = series.GetFirstChild<C.Values>()?
                        .GetFirstChild<C.NumberReference>()?
                        .GetFirstChild<C.NumberingCache>()?
                        .Elements<C.NumericPoint>()
                        .Select(point => double.Parse(point.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                        .ToArray() ?? Array.Empty<double>();
                    Assert.Equal(new[] { 12d, 18d, 15d }, values);

                    var marker = series.GetFirstChild<C.Marker>();
                    Assert.NotNull(marker);
                    Assert.Equal(C.MarkerStyleValues.Circle, marker!.Symbol?.Val?.Value);
                    Assert.Equal((byte)9, marker.Size?.Val?.Value);

                    string? markerFillColor = marker.GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("ED7D31", markerFillColor);

                    string? lineColor = series.GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("ED7D31", lineColor);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
