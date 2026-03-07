using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public class PowerPointChartFormatting {
        [Fact]
        public void CanFormatChartElements() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetTitle("Sales Trend")
                        .SetTitleTextStyle(fontSizePoints: 18, bold: true, color: "1F4E79", fontName: "Calibri")
                        .SetLegend(C.LegendPositionValues.Right)
                        .SetLegendTextStyle(fontSizePoints: 9, italic: true, color: "404040", fontName: "Calibri")
                        .SetDataLabels(showValue: true)
                        .SetDataLabelPosition(C.DataLabelPositionValues.OutsideEnd)
                        .SetDataLabelNumberFormat("#,##0.0", sourceLinked: false)
                        .SetCategoryAxisTitle("Quarter")
                        .SetValueAxisTitle("Revenue");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    C.Chart chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                    string? titleText = chart
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.Text>()?
                        .Text;
                    Assert.Equal("Sales Trend", titleText);

                    A.RunProperties? titleRunProps = chart
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.RunProperties>();
                    Assert.Equal(1800, titleRunProps?.FontSize?.Value);
                    Assert.True(titleRunProps?.Bold?.Value);
                    Assert.Equal("Calibri", titleRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", titleRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.LegendPositionValues? legendPosition = chart.GetFirstChild<C.Legend>()?.LegendPosition?.Val?.Value;
                    Assert.Equal(C.LegendPositionValues.Right, legendPosition);

                    A.DefaultRunProperties? legendRunProps = chart
                        .GetFirstChild<C.Legend>()?
                        .GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, legendRunProps?.FontSize?.Value);
                    Assert.True(legendRunProps?.Italic?.Value);
                    Assert.Equal("Calibri", legendRunProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("404040", legendRunProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.BarChart barChart = chart.PlotArea!.GetFirstChild<C.BarChart>()!;
                    C.DataLabels? dataLabels = barChart.GetFirstChild<C.DataLabels>();
                    bool? showValue = dataLabels?.GetFirstChild<C.ShowValue>()?.Val?.Value;
                    Assert.True(showValue);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, dataLabels?.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("#,##0.0", dataLabels?.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.False(dataLabels?.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value);

                    string? categoryTitle = chart.PlotArea!
                        .GetFirstChild<C.CategoryAxis>()?
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.Text>()?
                        .Text;
                    Assert.Equal("Quarter", categoryTitle);

                    string? valueTitle = chart.PlotArea!
                        .GetFirstChild<C.ValueAxis>()?
                        .GetFirstChild<C.Title>()?
                        .GetFirstChild<C.ChartText>()?
                        .GetFirstChild<C.RichText>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.Run>()?
                        .GetFirstChild<A.Text>()?
                        .Text;
                    Assert.Equal("Revenue", valueTitle);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanStyleChartAndPlotArea() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1.25)
                        .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "00B0F0", lineWidthPoints: 0.5);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.ChartSpace chartSpace = chartPart.ChartSpace;

                    C.ShapeProperties? chartProps = chartSpace.GetFirstChild<C.ShapeProperties>();
                    Assert.NotNull(chartProps);
                    A.SolidFill? chartFill = chartProps!.GetFirstChild<A.SolidFill>();
                    A.Outline? chartOutline = chartProps.GetFirstChild<A.Outline>();
                    Assert.Equal("F2F2F2", chartFill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal("404040", chartOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(1.25d * 12700d), chartOutline?.Width?.Value);

                    C.PlotArea plotArea = chartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                    C.ShapeProperties? plotProps = plotArea.GetFirstChild<C.ShapeProperties>();
                    Assert.NotNull(plotProps);
                    A.SolidFill? plotFill = plotProps!.GetFirstChild<A.SolidFill>();
                    A.Outline? plotOutline = plotProps.GetFirstChild<A.Outline>();
                    Assert.Equal("FFFFFF", plotFill?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal("00B0F0", plotOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.5d * 12700d), plotOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanStyleDataLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddChart();
                    chart.SetDataLabels(showValue: true)
                        .SetDataLabelTextStyle(fontSizePoints: 9, bold: true, color: "1F4E79", fontName: "Calibri")
                        .SetDataLabelShapeStyle(fillColor: "FFFFFF", lineColor: "1F4E79", lineWidthPoints: 0.5)
                        .SetDataLabelLeaderLines(showLeaderLines: true, lineColor: "1F4E79", lineWidthPoints: 0.5)
                        .SetDataLabelSeparator(" | ");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.BarChart>()!
                        .GetFirstChild<C.DataLabels>()!;

                    A.DefaultRunProperties? runProps = labels.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
                    Assert.NotNull(shapeProps);
                    Assert.Equal("FFFFFF", shapeProps!.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps.GetFirstChild<A.Outline>();
                    Assert.Equal("1F4E79", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.5d * 12700d), outline?.Width?.Value);

                    Assert.Equal(" | ", labels.GetFirstChild<C.Separator>()?.Text);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("1F4E79", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.5d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanApplyChartDataLabelTemplate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var template = new PowerPointChartDataLabelTemplate {
                    ShowValue = true,
                    ShowCategoryName = true,
                    Position = C.DataLabelPositionValues.OutsideEnd,
                    NumberFormat = "0.0",
                    Separator = " - ",
                    SourceLinked = false,
                    ShowLeaderLines = true,
                    LeaderLineColor = "C00000",
                    LeaderLineWidthPoints = 0.75,
                    FontSizePoints = 9,
                    Bold = true,
                    FontName = "Calibri",
                    TextColor = "1F4E79",
                    FillColor = "FFFFFF",
                    LineColor = "1F4E79",
                    LineWidthPoints = 0.5
                };

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddPieChart();
                    chart.SetDataLabelTemplate(template);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.PieChart>()!
                        .GetFirstChild<C.DataLabels>()!;

                    Assert.True(labels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.True(labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.0", labels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.False(labels.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value);
                    Assert.Equal(" - ", labels.GetFirstChild<C.Separator>()?.Text);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    A.DefaultRunProperties? runProps = labels.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
                    Assert.Equal("FFFFFF", shapeProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps?.GetFirstChild<A.Outline>();
                    Assert.Equal("1F4E79", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.5d * 12700d), outline?.Width?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanConfigureChartDataLabelCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "North", "South", "West" },
                    new[] { new PowerPointChartSeries("Sales", new[] { 42d, 31d, 27d }) });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddPieChart(data);
                    chart.SetDataLabelCallouts(enabled: true, lineColor: "C00000", lineWidthPoints: 0.75);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.PieChart>()!
                        .GetFirstChild<C.DataLabels>()!;

                    Assert.True(labels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.False(labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.False(labels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value);
                    Assert.False(labels.GetFirstChild<C.ShowPercent>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanApplySeriesDataLabelTemplate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "North", "South", "West" },
                    new[] { new PowerPointChartSeries("Sales", new[] { 42d, 31d, 27d }) });
                var template = new PowerPointChartDataLabelTemplate {
                    ShowValue = true,
                    ShowCategoryName = true,
                    Position = C.DataLabelPositionValues.OutsideEnd,
                    NumberFormat = "0.0",
                    Separator = " - ",
                    SourceLinked = false,
                    ShowLeaderLines = true,
                    LeaderLineColor = "C00000",
                    LeaderLineWidthPoints = 0.75,
                    FontSizePoints = 9,
                    Bold = true,
                    FontName = "Calibri",
                    TextColor = "1F4E79",
                    FillColor = "FFFFFF",
                    LineColor = "1F4E79",
                    LineWidthPoints = 0.5
                };

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddPieChart(data);
                    chart.SetSeriesDataLabelTemplate(0, template);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.PieChart>()!
                        .Elements<C.PieChartSeries>()
                        .Single()
                        .GetFirstChild<C.DataLabels>()!;

                    Assert.True(labels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.True(labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.0", labels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.False(labels.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value);
                    Assert.Equal(" - ", labels.GetFirstChild<C.Separator>()?.Text);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    A.DefaultRunProperties? runProps = labels.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(900, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("1F4E79", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
                    Assert.Equal("FFFFFF", shapeProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps?.GetFirstChild<A.Outline>();
                    Assert.Equal("1F4E79", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.5d * 12700d), outline?.Width?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanStyleSeriesDataLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "Jan", "Feb", "Mar" },
                    new[] {
                        new PowerPointChartSeries("Revenue", new[] { 10d, 12d, 14d }),
                        new PowerPointChartSeries("Forecast", new[] { 11d, 13d, 15d })
                    });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddLineChart(data);
                    chart.SetDataLabels(showValue: true)
                        .SetSeriesDataLabelTextStyle("Forecast", fontSizePoints: 11, bold: true, color: "C00000", fontName: "Calibri")
                        .SetSeriesDataLabelShapeStyle(1, fillColor: "FFF2CC", lineColor: "C00000", lineWidthPoints: 0.75)
                        .SetSeriesDataLabelLeaderLines("Forecast", showLeaderLines: true, lineColor: "C00000", lineWidthPoints: 0.75)
                        .SetSeriesDataLabelSeparator(1, " / ");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.LineChartSeries[] series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.LineChart>()!
                        .Elements<C.LineChartSeries>()
                        .ToArray();

                    Assert.Equal(2, series.Length);
                    Assert.Null(series[0].GetFirstChild<C.DataLabels>());

                    C.DataLabels labels = series[1].GetFirstChild<C.DataLabels>()!;
                    var children = series[1].ChildElements.ToList();
                    Assert.True(children.IndexOf(labels) < children.IndexOf(series[1].GetFirstChild<C.CategoryAxisData>()!));

                    A.DefaultRunProperties? runProps = labels.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(1100, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("C00000", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = labels.GetFirstChild<C.ChartShapeProperties>();
                    Assert.NotNull(shapeProps);
                    Assert.Equal("FFF2CC", shapeProps!.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps.GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), outline?.Width?.Value);

                    Assert.Equal(" / ", labels.GetFirstChild<C.Separator>()?.Text);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanConfigureSeriesDataLabelsAndCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "Jan", "Feb", "Mar" },
                    new[] {
                        new PowerPointChartSeries("Revenue", new[] { 10d, 12d, 14d }),
                        new PowerPointChartSeries("Forecast", new[] { 11d, 13d, 15d })
                    });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddLineChart(data);
                    chart.SetSeriesDataLabels("Forecast", showValue: true, showCategoryName: true, position: C.DataLabelPositionValues.Top,
                            numberFormat: "0.0", sourceLinked: false)
                        .SetSeriesDataLabelCallouts(0, enabled: true, lineColor: "C00000", lineWidthPoints: 0.75);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.LineChartSeries[] series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.LineChart>()!
                        .Elements<C.LineChartSeries>()
                        .ToArray();

                    Assert.Equal(2, series.Length);

                    C.DataLabels revenueLabels = series[0].GetFirstChild<C.DataLabels>()!;
                    Assert.True(revenueLabels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.False(revenueLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, revenueLabels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.True(revenueLabels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);
                    A.Outline? revenueLeaderOutline = revenueLabels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", revenueLeaderOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), revenueLeaderOutline?.Width?.Value);

                    C.DataLabels forecastLabels = series[1].GetFirstChild<C.DataLabels>()!;
                    Assert.True(forecastLabels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.True(forecastLabels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.False(forecastLabels.GetFirstChild<C.ShowSeriesName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.Top, forecastLabels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.0", forecastLabels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.False(forecastLabels.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanConfigurePointLevelDataLabelOverrides() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "North", "South", "West" },
                    new[] { new PowerPointChartSeries("Sales", new[] { 42d, 31d, 27d }) });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddPieChart(data);
                    chart.SetSeriesDataLabels(0, showValue: true, position: C.DataLabelPositionValues.BestFit, numberFormat: "0.0", sourceLinked: false)
                        .SetSeriesDataLabelSeparator(0, " - ")
                        .SetSeriesDataLabelLeaderLines(0, showLeaderLines: true, lineColor: "C00000", lineWidthPoints: 0.75)
                        .SetSeriesDataLabelCalloutsForPoint(0, 1, enabled: true)
                        .SetSeriesDataLabelForPoint(0, 1, showValue: true, showCategoryName: true, position: C.DataLabelPositionValues.OutsideEnd,
                            numberFormat: "0.00", sourceLinked: false)
                        .SetSeriesDataLabelSeparatorForPoint(0, 1, " | ")
                        .SetSeriesDataLabelTextStyleForPoint(0, 1, fontSizePoints: 11, bold: true, color: "C00000", fontName: "Calibri")
                        .SetSeriesDataLabelShapeStyleForPoint(0, 1, fillColor: "FFF2CC", lineColor: "C00000", lineWidthPoints: 0.75);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.PieChart>()!
                        .Elements<C.PieChartSeries>()
                        .Single()
                        .GetFirstChild<C.DataLabels>()!;

                    Assert.True(labels.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.BestFit, labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.0", labels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.Equal(" - ", labels.GetFirstChild<C.Separator>()?.Text);
                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);

                    C.DataLabel? pointLabel = labels.Elements<C.DataLabel>()
                        .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == 1U);
                    Assert.NotNull(pointLabel);
                    Assert.True(pointLabel!.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.True(pointLabel.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, pointLabel.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.00", pointLabel.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.Equal(" | ", pointLabel.GetFirstChild<C.Separator>()?.Text);

                    A.DefaultRunProperties? runProps = pointLabel.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(1100, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("C00000", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = pointLabel.GetFirstChild<C.ChartShapeProperties>();
                    Assert.Equal("FFF2CC", shapeProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps?.GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), outline?.Width?.Value);

                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanApplyPointLevelDataLabelTemplate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "North", "South", "West" },
                    new[] { new PowerPointChartSeries("Sales", new[] { 42d, 31d, 27d }) });
                var template = new PowerPointChartDataLabelTemplate {
                    ShowValue = true,
                    ShowCategoryName = true,
                    Position = C.DataLabelPositionValues.OutsideEnd,
                    NumberFormat = "0.00",
                    Separator = " | ",
                    SourceLinked = false,
                    ShowLeaderLines = true,
                    LeaderLineColor = "C00000",
                    LeaderLineWidthPoints = 0.75,
                    FontSizePoints = 11,
                    Bold = true,
                    FontName = "Calibri",
                    TextColor = "C00000",
                    FillColor = "FFF2CC",
                    LineColor = "C00000",
                    LineWidthPoints = 0.75
                };

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddPieChart(data);
                    chart.SetSeriesDataLabelTemplateForPoint(0, 1, template);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.DataLabels labels = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.PieChart>()!
                        .Elements<C.PieChartSeries>()
                        .Single()
                        .GetFirstChild<C.DataLabels>()!;

                    Assert.True(labels.GetFirstChild<C.ShowLeaderLines>()?.Val?.Value);
                    A.Outline? leaderLineOutline = labels.GetFirstChild<C.LeaderLines>()?
                        .GetFirstChild<C.ChartShapeProperties>()?
                        .GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", leaderLineOutline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), leaderLineOutline?.Width?.Value);

                    C.DataLabel? pointLabel = labels.Elements<C.DataLabel>()
                        .FirstOrDefault(item => item.GetFirstChild<C.Index>()?.Val?.Value == 1U);
                    Assert.NotNull(pointLabel);
                    Assert.True(pointLabel!.GetFirstChild<C.ShowValue>()?.Val?.Value);
                    Assert.True(pointLabel.GetFirstChild<C.ShowCategoryName>()?.Val?.Value);
                    Assert.Equal(C.DataLabelPositionValues.OutsideEnd, pointLabel.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                    Assert.Equal("0.00", pointLabel.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                    Assert.False(pointLabel.GetFirstChild<C.NumberingFormat>()?.SourceLinked?.Value);
                    Assert.Equal(" | ", pointLabel.GetFirstChild<C.Separator>()?.Text);

                    A.DefaultRunProperties? runProps = pointLabel.GetFirstChild<C.TextProperties>()?
                        .GetFirstChild<A.Paragraph>()?
                        .GetFirstChild<A.ParagraphProperties>()?
                        .GetFirstChild<A.DefaultRunProperties>();
                    Assert.Equal(1100, runProps?.FontSize?.Value);
                    Assert.True(runProps?.Bold?.Value);
                    Assert.Equal("Calibri", runProps?.GetFirstChild<A.LatinFont>()?.Typeface?.Value);
                    Assert.Equal("C00000", runProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);

                    C.ChartShapeProperties? shapeProps = pointLabel.GetFirstChild<C.ChartShapeProperties>();
                    Assert.Equal("FFF2CC", shapeProps?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    A.Outline? outline = shapeProps?.GetFirstChild<A.Outline>();
                    Assert.Equal("C00000", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(0.75d * 12700d), outline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanAddSeriesTrendlineToLineChart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointChartData(
                    new[] { "Jan", "Feb", "Mar" },
                    new[] { new PowerPointChartSeries("Revenue", new[] { 12d, 18d, 15d }) });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddLineChart(data);
                    chart.SetSeriesTrendline(0, C.TrendlineValues.Polynomial, order: 2, forward: 1.5, backward: 0.5,
                            intercept: 10, displayEquation: true, displayRSquared: true, lineColor: "ED7D31", lineWidthPoints: 1.5)
                        .SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 7, fillColor: "FFFFFF", lineColor: "ED7D31");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.LineChartSeries series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.LineChart>()!
                        .Elements<C.LineChartSeries>()
                        .Single();

                    C.Marker? marker = series.GetFirstChild<C.Marker>();
                    C.Trendline? trendline = series.GetFirstChild<C.Trendline>();
                    Assert.NotNull(marker);
                    Assert.NotNull(trendline);
                    var children = series.ChildElements.ToList();
                    Assert.True(children.IndexOf(marker!) < children.IndexOf(trendline!));

                    Assert.Equal(C.TrendlineValues.Polynomial, trendline!.GetFirstChild<C.TrendlineType>()?.Val?.Value);
                    Assert.Equal((byte)2, trendline.GetFirstChild<C.PolynomialOrder>()?.Val?.Value);
                    Assert.Equal(1.5d, trendline.GetFirstChild<C.Forward>()?.Val?.Value);
                    Assert.Equal(0.5d, trendline.GetFirstChild<C.Backward>()?.Val?.Value);
                    Assert.Equal(10d, trendline.GetFirstChild<C.Intercept>()?.Val?.Value);
                    Assert.True(trendline.GetFirstChild<C.DisplayEquation>()?.Val?.Value);
                    Assert.True(trendline.GetFirstChild<C.DisplayRSquaredValue>()?.Val?.Value);

                    A.Outline? outline = trendline.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<A.Outline>();
                    Assert.Equal("ED7D31", outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
                    Assert.Equal((int)Math.Round(1.5d * 12700d), outline?.Width?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanClearSeriesTrendlineByName() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                var data = new PowerPointScatterChartData(new[] {
                    new PowerPointScatterChartSeries("Revenue", new[] { 1d, 2d, 3d }, new[] { 10d, 15d, 12d })
                });

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart chart = slide.AddScatterChart(data);
                    chart.SetSeriesTrendline("Revenue", C.TrendlineValues.Linear, lineColor: "5B9BD5", lineWidthPoints: 2)
                        .ClearSeriesTrendline("Revenue");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.First().ChartParts.First();
                    OpenXmlValidator validator = new OpenXmlValidator();
                    Assert.Empty(validator.Validate(chartPart.ChartSpace));

                    C.ScatterChartSeries series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                        .GetFirstChild<C.PlotArea>()!
                        .GetFirstChild<C.ScatterChart>()!
                        .Elements<C.ScatterChartSeries>()
                        .Single();

                    Assert.Null(series.GetFirstChild<C.Trendline>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
