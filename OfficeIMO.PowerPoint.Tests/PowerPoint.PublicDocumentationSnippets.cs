using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPublicDocumentationSnippetsTests {
        [Fact]
        public void ProductPageQuickStart_UsesTheSupportedPublicApi() {
            string filePath = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointSlide intro = presentation.AddSlide();
                    intro.AddTitle("Product overview",global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(1.2),global::OfficeIMO.OfficeLength.Centimeters(22),global::OfficeIMO.OfficeLength.Centimeters(1.4));
                    PowerPointTextBox highlights = intro.AddTextBox(string.Empty,global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(3.0),global::OfficeIMO.OfficeLength.Centimeters(12),global::OfficeIMO.OfficeLength.Centimeters(5.5));
                    highlights.AddBullets(new[] {
                        "Revenue grew 18% year over year",
                        "Customer satisfaction reached 94%",
                        "Delivery expanded to 12 markets"
                    });

                    var data = new OfficeChartData(
                        new[] { "Q1", "Q2", "Q3", "Q4" },
                        new[] { new OfficeSeries("Revenue", new[] { 3.2, 3.8, 4.1, 4.9 }) });

                    PowerPointSlide chartSlide = presentation.AddSlide();
                    chartSlide.AddTitle("Revenue by quarter",global::OfficeIMO.OfficeLength.Centimeters(1.5),global::OfficeIMO.OfficeLength.Centimeters(1.2),global::OfficeIMO.OfficeLength.Centimeters(22),global::OfficeIMO.OfficeLength.Centimeters(1.4));
                    chartSlide.AddChart(OfficeChartType.ColumnClustered, data, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Centimeters(1.5), Top = global::OfficeIMO.OfficeLength.Centimeters(3.0), Width = global::OfficeIMO.OfficeLength.Centimeters(22), Height = global::OfficeIMO.OfficeLength.Centimeters(9), Accessibility = new PowerPointChartAccessibilityOptions {
                                AlternativeText = "Quarterly revenue increased from 3.2 to 4.9"
                            } })
                        .SetTitle("2025 revenue")
                        .SetLegend(OfficeChartLegendPosition.Bottom);

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using PowerPointDocument reopened = PowerPointDocument.Load(filePath);
                Assert.Equal(2, reopened.Slides.Count);
                Assert.Single(reopened.Slides[1].Charts);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
