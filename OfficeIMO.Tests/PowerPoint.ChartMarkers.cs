using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointChartMarkerTests {
        [Fact]
        public void CanSetSeriesMarkers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    ConvertBarChartToLine(chartPart);
                    chartPart.ChartSpace?.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointChart chart = presentation.Slides.First().Charts.First();
                    chart.SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 8, fillColor: "FF0000", lineColor: "00FF00", lineWidthPoints: 1);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.LineChartSeries series = chartPart.ChartSpace!
                        .GetFirstChild<C.Chart>()!
                        .PlotArea!
                        .GetFirstChild<C.LineChart>()!
                        .Elements<C.LineChartSeries>()
                        .First();

                    C.Marker marker = series.GetFirstChild<C.Marker>()!;
                    Assert.Equal(C.MarkerStyleValues.Circle, marker.Symbol?.Val?.Value);
                    Assert.Equal((byte)8, marker.Size?.Val?.Value);

                    string? fill = marker.ChartShapeProperties?
                        .GetFirstChild<A.SolidFill>()?
                        .GetFirstChild<A.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("FF0000", fill);

                    A.Outline? outline = marker.ChartShapeProperties?.GetFirstChild<A.Outline>();
                    string? line = outline?
                        .GetFirstChild<A.SolidFill>()?
                        .GetFirstChild<A.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("00FF00", line);
                    Assert.Equal(12700, outline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void ConvertBarChartToLine(ChartPart chartPart) {
            C.Chart? chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.PlotArea;
            if (plotArea == null) {
                return;
            }

            C.BarChart? barChart = plotArea.GetFirstChild<C.BarChart>();
            if (barChart == null) {
                return;
            }

            var axisIds = barChart.Elements<C.AxisId>()
                .Select(id => id.Val?.Value)
                .Where(id => id.HasValue)
                .Select(id => id!.Value)
                .ToList();

            C.LineChart lineChart = new C.LineChart(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false });

            foreach (C.BarChartSeries barSeries in barChart.Elements<C.BarChartSeries>()) {
                C.LineChartSeries lineSeries = new C.LineChartSeries();
                if (barSeries.Index != null) {
                    lineSeries.Append((C.Index)barSeries.Index.CloneNode(true));
                }
                if (barSeries.Order != null) {
                    lineSeries.Append((C.Order)barSeries.Order.CloneNode(true));
                }
                C.SeriesText? seriesText = barSeries.GetFirstChild<C.SeriesText>();
                if (seriesText != null) {
                    lineSeries.Append((C.SeriesText)seriesText.CloneNode(true));
                }
                C.CategoryAxisData? categoryAxisData = barSeries.GetFirstChild<C.CategoryAxisData>();
                if (categoryAxisData != null) {
                    lineSeries.Append((C.CategoryAxisData)categoryAxisData.CloneNode(true));
                }
                C.Values? values = barSeries.GetFirstChild<C.Values>();
                if (values != null) {
                    lineSeries.Append((C.Values)values.CloneNode(true));
                }
                lineChart.Append(lineSeries);
            }

            foreach (uint axisId in axisIds) {
                lineChart.Append(new C.AxisId { Val = axisId });
            }

            barChart.InsertAfterSelf(lineChart);
            barChart.Remove();
        }
    }
}
