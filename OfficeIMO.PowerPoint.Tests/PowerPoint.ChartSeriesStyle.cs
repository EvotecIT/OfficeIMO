using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointChartSeriesStyleTests {
        [Fact]
        public void CanStyleSeriesFillAndLine() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetSeriesFillColor(1, "FF0000")
                        .SetSeriesLineColor(0, "00FF00", widthPoints: 1);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.BarChart barChart = chart.PlotArea!.GetFirstChild<C.BarChart>()!;

                    C.BarChartSeries series0 = barChart.Elements<C.BarChartSeries>().ElementAt(0);
                    C.BarChartSeries series1 = barChart.Elements<C.BarChartSeries>().ElementAt(1);

                    string? fill = series1.ChartShapeProperties?
                        .GetFirstChild<A.SolidFill>()?
                        .GetFirstChild<A.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("FF0000", fill);

                    A.Outline? outline = series0.ChartShapeProperties?.GetFirstChild<A.Outline>();
                    string? lineColor = outline?
                        .GetFirstChild<A.SolidFill>()?
                        .GetFirstChild<A.RgbColorModelHex>()?
                        .Val?.Value;
                    Assert.Equal("00FF00", lineColor);
                    Assert.Equal(12700, outline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
