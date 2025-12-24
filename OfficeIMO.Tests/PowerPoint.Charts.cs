using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointCharts {
        [Fact]
        public void CanCreateMultipleChartsWithUniqueParts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] { new PowerPointChartSeries("Series 1", new[] { 1d, 2d, 3d }) });

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddChart(data);
                slide.AddChart(data);
                presentation.Save();
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                var chartParts = document.PresentationPart!
                    .SlideParts
                    .SelectMany(slidePart => slidePart.GetPartsOfType<ChartPart>())
                    .ToList();
                Assert.Equal(2, chartParts.Count);
                Assert.Equal(2, chartParts.Select(p => p.Uri.ToString()).Distinct().Count());

                foreach (ChartPart chartPart in chartParts) {
                    ChartStylePart? stylePart = chartPart.GetPartsOfType<ChartStylePart>().FirstOrDefault();
                    ChartColorStylePart? colorPart = chartPart.GetPartsOfType<ChartColorStylePart>().FirstOrDefault();
                    EmbeddedPackagePart? embeddedPart = chartPart.GetPartsOfType<EmbeddedPackagePart>().FirstOrDefault();

                    Assert.NotNull(stylePart);
                    Assert.NotNull(colorPart);
                    Assert.NotNull(embeddedPart);

                    Assert.StartsWith("/ppt/charts/chart", chartPart.Uri.ToString(), StringComparison.OrdinalIgnoreCase);
                    Assert.StartsWith("/ppt/charts/style", stylePart!.Uri.ToString(), StringComparison.OrdinalIgnoreCase);
                    Assert.StartsWith("/ppt/charts/colors", colorPart!.Uri.ToString(), StringComparison.OrdinalIgnoreCase);
                    Assert.StartsWith("/ppt/embeddings/Microsoft_Excel_Worksheet", embeddedPart!.Uri.ToString(), StringComparison.OrdinalIgnoreCase);

                    C.ExternalData? externalData = chartPart.ChartSpace?
                        .Descendants<C.ExternalData>()
                        .FirstOrDefault();
                    string? relId = externalData?.Id?.Value;
                    Assert.False(string.IsNullOrWhiteSpace(relId));
                    Assert.Equal(relId, chartPart.GetIdOfPart(embeddedPart!));
                }
            }

            File.Delete(filePath);
        }
    }
}
