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
            var updatedData = new PowerPointChartData(
                new[] { "Central", "East", "West" },
                new[] { new PowerPointChartSeries("Updated Revenue", new[] { 15d, 25d, 35d }) });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = doughnut
                        ? slide.AddDoughnutChartInches(data, 1, 1, 4, 3)
                        : slide.AddPieChartPoints(data, 72, 72, 288, 216);

                    chart.SetTitle(doughnut ? "Revenue Doughnut" : "Revenue Pie")
                        .SetDataLabels(showValue: true, showPercent: true)
                        .SetDataLabelPosition(C.DataLabelPositionValues.BestFit)
                        .SetDataLabelNumberFormat("0.0%", sourceLinked: false)
                        .SetSeriesFillColor(0, "4472C4");
                    chart.UpdateData(updatedData);

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
                    Assert.Equal(new[] { "Central", "East", "West" }, categories);

                    double[] values = series.GetFirstChild<C.Values>()?
                        .GetFirstChild<C.NumberReference>()?
                        .GetFirstChild<C.NumberingCache>()?
                        .Elements<C.NumericPoint>()
                        .Select(point => double.Parse(point.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                        .ToArray() ?? Array.Empty<double>();
                    Assert.Equal(new[] { 15d, 25d, 35d }, values);

                    bool? showPercent = chartElement.GetFirstChild<C.DataLabels>()?
                        .GetFirstChild<C.ShowPercent>()?
                        .Val?.Value;
                    Assert.True(showPercent);
                    Assert.Equal(C.DataLabelPositionValues.BestFit,
                        chartElement.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.0%",
                        chartElement.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);

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

        [Fact]
        public void CanCreateAndUpdateScatterCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var initialData = new PowerPointScatterChartData(new[] {
                new PowerPointScatterChartSeries("Revenue", new[] { 1d, 2d, 3d }, new[] { 10d, 15d, 12d })
            });

            var updatedData = new PowerPointScatterChartData(new[] {
                new PowerPointScatterChartSeries("Revenue", new[] { 1d, 2d, 3d, 4d }, new[] { 10d, 15d, 12d, 18d }),
                new PowerPointScatterChartSeries("Forecast", new[] { 1.5d, 2.5d, 3.5d }, new[] { 11d, 14d, 17d })
            });

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart(initialData);

                    chart.UpdateData(updatedData)
                        .SetTitle("Revenue Scatter")
                        .SetDataLabels(showValue: true)
                        .SetDataLabelPosition(C.DataLabelPositionValues.Right)
                        .SetDataLabelNumberFormat("0.00", sourceLinked: false)
                        .SetSeriesLineColor("Revenue", "5B9BD5", 2)
                        .SetSeriesMarker("Forecast", C.MarkerStyleValues.Diamond, size: 8, fillColor: "ED7D31", lineColor: "A64D13");

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace!.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    Assert.Null(plotArea.GetFirstChild<C.BarChart>());
                    Assert.Null(plotArea.GetFirstChild<C.LineChart>());
                    Assert.Null(plotArea.GetFirstChild<C.PieChart>());
                    Assert.Null(plotArea.GetFirstChild<C.DoughnutChart>());
                    Assert.Empty(plotArea.Elements<C.CategoryAxis>());

                    C.ValueAxis[] axes = plotArea.Elements<C.ValueAxis>().ToArray();
                    Assert.Equal(2, axes.Length);
                    Assert.Contains(axes, axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
                    Assert.Contains(axes, axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);

                    C.ScatterChart scatterChart = plotArea.GetFirstChild<C.ScatterChart>()!;
                    Assert.Equal(C.ScatterStyleValues.LineMarker, scatterChart.ScatterStyle?.Val?.Value);

                    C.ScatterChartSeries[] series = scatterChart.Elements<C.ScatterChartSeries>().ToArray();
                    Assert.Equal(2, series.Length);

                    double[] revenueX = series[0].GetFirstChild<C.XValues>()?
                        .GetFirstChild<C.NumberReference>()?
                        .GetFirstChild<C.NumberingCache>()?
                        .Elements<C.NumericPoint>()
                        .Select(point => double.Parse(point.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                        .ToArray() ?? Array.Empty<double>();
                    Assert.Equal(new[] { 1d, 2d, 3d, 4d }, revenueX);

                    double[] forecastY = series[1].GetFirstChild<C.YValues>()?
                        .GetFirstChild<C.NumberReference>()?
                        .GetFirstChild<C.NumberingCache>()?
                        .Elements<C.NumericPoint>()
                        .Select(point => double.Parse(point.NumericValue?.Text ?? "0", CultureInfo.InvariantCulture))
                        .ToArray() ?? Array.Empty<double>();
                    Assert.Equal(new[] { 11d, 14d, 17d }, forecastY);

                    bool? showValue = scatterChart.GetFirstChild<C.DataLabels>()?
                        .GetFirstChild<C.ShowValue>()?
                        .Val?.Value;
                    Assert.True(showValue);
                    Assert.Equal(C.DataLabelPositionValues.Right,
                        scatterChart.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.00",
                        scatterChart.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);

                    string? lineColor = series[0].GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("5B9BD5", lineColor);

                    C.Marker? marker = series[1].GetFirstChild<C.Marker>();
                    Assert.NotNull(marker);
                    Assert.Equal(C.MarkerStyleValues.Diamond, marker!.Symbol?.Val?.Value);

                    string? markerFillColor = marker.GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?
                        .GetFirstChild<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("ED7D31", markerFillColor);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
