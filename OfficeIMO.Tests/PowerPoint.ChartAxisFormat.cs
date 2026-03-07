using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
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

        [Fact]
        public void CanStyleAxisTitles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetCategoryAxisTitle("Quarter")
                        .SetCategoryAxisTitleTextStyle(fontSizePoints: 11, bold: true, color: "1F4E79", fontName: "Calibri")
                        .SetValueAxisTitle("Revenue")
                        .SetValueAxisTitleTextStyle(fontSizePoints: 10, italic: true, color: "C55A11", fontName: "Arial");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                    A.RunProperties? categoryRunProps = chart.PlotArea!
                        .GetFirstChild<C.CategoryAxis>()?
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.RunProperties>();
                    Assert.Equal(1100, categoryRunProps?.FontSize?.Value);
                    Assert.True(categoryRunProps?.Bold?.Value);
                    Assert.Equal("Calibri", categoryRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", categoryRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    A.RunProperties? valueRunProps = chart.PlotArea!
                        .GetFirstChild<C.ValueAxis>()?
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.RunProperties>();
                    Assert.Equal(1000, valueRunProps?.FontSize?.Value);
                    Assert.True(valueRunProps?.Italic?.Value);
                    Assert.Equal("Arial", valueRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("C55A11", valueRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetAxisScaleAndCrossing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetCategoryAxisReverseOrder()
                        .SetValueAxisScale(minimum: 0, maximum: 100, majorUnit: 25, minorUnit: 5, reverseOrder: true)
                        .SetValueAxisCrossing(C.CrossesValues.Maximum)
                        .SetCategoryAxisCrossing(C.CrossesValues.Minimum);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    C.CategoryAxis categoryAxis = plotArea.GetFirstChild<C.CategoryAxis>()!;
                    C.Scaling? categoryScaling = categoryAxis.GetFirstChild<C.Scaling>();
                    Assert.Equal(C.OrientationValues.MaxMin, categoryScaling?.GetFirstChild<C.Orientation>()?.Val?.Value);
                    Assert.Equal(C.CrossesValues.Minimum, categoryAxis.GetFirstChild<C.Crosses>()?.Val?.Value);

                    C.ValueAxis valueAxis = plotArea.GetFirstChild<C.ValueAxis>()!;
                    C.Scaling? valueScaling = valueAxis.GetFirstChild<C.Scaling>();
                    Assert.Equal(0d, (double?)valueScaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                    Assert.Equal(100d, (double?)valueScaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                    Assert.Equal(C.OrientationValues.MaxMin, valueScaling?.GetFirstChild<C.Orientation>()?.Val?.Value);
                    Assert.Equal(25d, (double?)valueAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);
                    Assert.Equal(5d, (double?)valueAxis.GetFirstChild<C.MinorUnit>()?.Val?.Value);
                    Assert.Equal(C.CrossesValues.Maximum, valueAxis.GetFirstChild<C.Crosses>()?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetValueAxisCrossingAtValue() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetValueAxisCrossing(C.CrossesValues.AutoZero, crossesAt: 2.5);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.ValueAxis valueAxis = chart.PlotArea!.GetFirstChild<C.ValueAxis>()!;

                    Assert.Equal(2.5d, (double?)valueAxis.GetFirstChild<C.CrossesAt>()?.Val?.Value);
                    Assert.Null(valueAxis.GetFirstChild<C.Crosses>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
