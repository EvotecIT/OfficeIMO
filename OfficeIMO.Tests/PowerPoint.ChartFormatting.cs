using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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
    }
}
