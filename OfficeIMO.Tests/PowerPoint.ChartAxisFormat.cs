using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
        public void CanStyleAxisLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetCategoryAxisLabelTextStyle(fontSizePoints: 9, bold: true, color: "404040", fontName: "Calibri")
                        .SetValueAxisLabelTextStyle(fontSizePoints: 10, italic: true, color: "1F4E79", fontName: "Arial");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                    A.DefaultRunProperties? categoryRunProps = chart.PlotArea!
                        .GetFirstChild<C.CategoryAxis>()?
                        .GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, categoryRunProps?.FontSize?.Value);
                    Assert.True(categoryRunProps?.Bold?.Value);
                    Assert.Equal("Calibri", categoryRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("404040", categoryRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    A.DefaultRunProperties? valueRunProps = chart.PlotArea!
                        .GetFirstChild<C.ValueAxis>()?
                        .GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(1000, valueRunProps?.FontSize?.Value);
                    Assert.True(valueRunProps?.Italic?.Value);
                    Assert.Equal("Arial", valueRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", valueRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
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

        [Fact]
        public void CanStyleScatterAxisLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterXAxisLabelTextStyle(fontSizePoints: 9, bold: true, color: "404040", fontName: "Calibri")
                        .SetScatterYAxisLabelTextStyle(fontSizePoints: 10, italic: true, color: "1F4E79", fontName: "Arial");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.ValueAxis[] axes = chart.PlotArea!
                        .Elements<C.ValueAxis>()
                        .ToArray();
                    Assert.Equal(2, axes.Length);

                    C.ValueAxis xAxis = axes.Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
                    A.DefaultRunProperties? xAxisRunProps = xAxis
                        .GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, xAxisRunProps?.FontSize?.Value);
                    Assert.True(xAxisRunProps?.Bold?.Value);
                    Assert.Equal("Calibri", xAxisRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("404040", xAxisRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ValueAxis yAxis = axes.Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
                    A.DefaultRunProperties? yAxisRunProps = yAxis
                        .GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(1000, yAxisRunProps?.FontSize?.Value);
                    Assert.True(yAxisRunProps?.Italic?.Value);
                    Assert.Equal("Arial", yAxisRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", yAxisRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetScatterAxisTitlesAndFormats() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterXAxisTitle("Month")
                        .SetScatterYAxisTitle("Revenue")
                        .SetScatterXAxisNumberFormat("0.0")
                        .SetScatterYAxisNumberFormat("#,##0.00");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                    C.ValueAxis[] axes = chart.PlotArea!
                        .Elements<C.ValueAxis>()
                        .ToArray();
                    Assert.Equal(2, axes.Length);

                    C.ValueAxis xAxis = axes.Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
                    C.ValueAxis yAxis = axes.Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);

                    string? xAxisTitle = xAxis.GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.Text>()?
                        .Text;
                    Assert.Equal("Month", xAxisTitle);

                    string? yAxisTitle = yAxis.GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.Text>()?
                        .Text;
                    Assert.Equal("Revenue", yAxisTitle);

                    C.NumberingFormat? xAxisFormat = xAxis.GetFirstChild<C.NumberingFormat>();
                    Assert.NotNull(xAxisFormat);
                    Assert.Equal("0.0", xAxisFormat!.FormatCode!.Value);
                    Assert.False(xAxisFormat.SourceLinked!.Value);

                    C.NumberingFormat? yAxisFormat = yAxis.GetFirstChild<C.NumberingFormat>();
                    Assert.NotNull(yAxisFormat);
                    Assert.Equal("#,##0.00", yAxisFormat!.FormatCode!.Value);
                    Assert.False(yAxisFormat.SourceLinked!.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ScatterAxisTitlesAndFormats_AreIgnoredForNonScatterCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetScatterXAxisTitle("Ignored X")
                        .SetScatterYAxisTitle("Ignored Y")
                        .SetScatterXAxisNumberFormat("0.0")
                        .SetScatterYAxisNumberFormat("#,##0.00");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    Assert.Null(plotArea.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.Title>());

                    C.ValueAxis valueAxis = plotArea.GetFirstChild<C.ValueAxis>()!;
                    Assert.Null(valueAxis.GetFirstChild<C.Title>());
                    C.NumberingFormat? valueAxisFormat = valueAxis.GetFirstChild<C.NumberingFormat>();
                    Assert.False(valueAxisFormat?.FormatCode?.Value == "#,##0.00" && valueAxisFormat.SourceLinked?.Value == false);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetScatterAxisScale() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterXAxisScale(minimum: 1, maximum: 10, majorUnit: 1, logScale: true)
                        .SetScatterYAxisScale(minimum: 0, maximum: 6, majorUnit: 1);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    C.ValueAxis xAxis = plotArea.Elements<C.ValueAxis>()
                        .Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
                    C.Scaling? xScaling = xAxis.GetFirstChild<C.Scaling>();
                    Assert.Equal(1d, (double?)xScaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                    Assert.Equal(10d, (double?)xScaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                    Assert.Equal(10d, (double?)xScaling?.GetFirstChild<C.LogBase>()?.Val?.Value);
                    Assert.Equal(1d, (double?)xAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);

                    C.ValueAxis yAxis = plotArea.Elements<C.ValueAxis>()
                        .Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
                    C.Scaling? yScaling = yAxis.GetFirstChild<C.Scaling>();
                    Assert.Equal(0d, (double?)yScaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                    Assert.Equal(6d, (double?)yScaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                    Assert.Equal(1d, (double?)yAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ScatterAxisScale_RejectsNonFiniteValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();

                    Assert.Throws<ArgumentOutOfRangeException>(() =>
                        chart.SetScatterXAxisScale(minimum: double.PositiveInfinity));
                    Assert.Throws<ArgumentOutOfRangeException>(() =>
                        chart.SetScatterXAxisScale(majorUnit: double.NaN));
                    Assert.Throws<ArgumentOutOfRangeException>(() =>
                        chart.SetScatterYAxisScale(minorUnit: double.PositiveInfinity));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ScatterAxisScale_RejectsContradictorySequentialBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterXAxisScale(maximum: 5);

                    Assert.Throws<ArgumentException>(() =>
                        chart.SetScatterXAxisScale(minimum: 10));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ScatterXAxisCrossing_RejectsNonPositiveOnLogScale() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterXAxisScale(minimum: 1, maximum: 10, logScale: true);

                    Assert.Throws<ArgumentOutOfRangeException>(() =>
                        chart.SetScatterXAxisCrossing(C.CrossesValues.AutoZero, crossesAt: 0));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetScatterYAxisCrossing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart();
                    chart.SetScatterYAxisCrossing(crossesAt: 2d);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                    C.PlotArea plotArea = chart.PlotArea!;

                    C.ValueAxis yAxis = plotArea.Elements<C.ValueAxis>()
                        .Single(axis => axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
                    Assert.Equal(2d, (double?)yAxis.GetFirstChild<C.CrossesAt>()?.Val?.Value);
                    Assert.Null(yAxis.GetFirstChild<C.Crosses>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
