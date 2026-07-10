using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPublicDocumentationSnippetsTests {
        [Fact]
        public void ProductPageQuickStart_UsesTheSupportedPublicApi() {
            string filePath = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointSlide intro = presentation.AddSlide();
                    intro.AddTitleCm("Product overview", 1.5, 1.2, 22, 1.4);
                    PowerPointTextBox highlights = intro.AddTextBoxCm(string.Empty, 1.5, 3.0, 12, 5.5);
                    highlights.AddBullets(new[] {
                        "Revenue grew 18% year over year",
                        "Customer satisfaction reached 94%",
                        "Delivery expanded to 12 markets"
                    });

                    var data = new PowerPointChartData(
                        new[] { "Q1", "Q2", "Q3", "Q4" },
                        new[] { new PowerPointChartSeries("Revenue", new[] { 3.2, 3.8, 4.1, 4.9 }) });

                    PowerPointSlide chartSlide = presentation.AddSlide();
                    chartSlide.AddTitleCm("Revenue by quarter", 1.5, 1.2, 22, 1.4);
                    chartSlide.AddChartCm(data, 1.5, 3.0, 22, 9)
                        .SetTitle("2025 revenue")
                        .SetLegend(LegendPositionValues.Bottom);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Open(filePath);
                Assert.Equal(2, reopened.Slides.Count);
                Assert.Single(reopened.Slides[1].Charts);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
