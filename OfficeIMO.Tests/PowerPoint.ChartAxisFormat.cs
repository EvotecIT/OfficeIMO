using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointChartAxisFormatTests {
        [Fact]
        public void CanSetAxisNumberFormats() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetCategoryAxisNumberFormat("0")
                        .SetValueAxisNumberFormat("#,##0.00");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                    C.NumberingFormat? categoryFormat = chart.PlotArea!
                        .GetFirstChild<C.CategoryAxis>()?
                        .GetFirstChild<C.NumberingFormat>();
                    Assert.NotNull(categoryFormat);
                    Assert.Equal("0", categoryFormat!.FormatCode!.Value);
                    Assert.False(categoryFormat.SourceLinked!.Value);

                    C.NumberingFormat? valueFormat = chart.PlotArea!
                        .GetFirstChild<C.ValueAxis>()?
                        .GetFirstChild<C.NumberingFormat>();
                    Assert.NotNull(valueFormat);
                    Assert.Equal("#,##0.00", valueFormat!.FormatCode!.Value);
                    Assert.False(valueFormat.SourceLinked!.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
