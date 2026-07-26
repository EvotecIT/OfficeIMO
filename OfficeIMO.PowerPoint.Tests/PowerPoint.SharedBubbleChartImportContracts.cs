using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class PowerPointSharedBubbleChartImportContractTests {
    [Fact]
    public void BubbleChart_RejectsInheritedSeriesOutline() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        C.BubbleChartSeries series = presentation.Slides[0].SlidePart
            .ChartParts.Single().ChartSpace!
            .Descendants<C.BubbleChartSeries>().Single();

        series.GetFirstChild<C.ChartShapeProperties>()!
            .RemoveAllChildren<A.Outline>();

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void BubbleChart_PreservesMultilineChartAndAxisTitles() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        chart.SetTitle("Chart title").SetScatterXAxisTitle("Axis title");
        C.Chart nativeChart = presentation.Slides[0].SlidePart.ChartParts
            .Single().ChartSpace!.GetFirstChild<C.Chart>()!;
        SetMultilineText(nativeChart.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!, "Risk", "Return", "Portfolio");
        C.ValueAxis horizontalAxis = nativeChart.GetFirstChild<C.PlotArea>()!
            .Elements<C.ValueAxis>().Single(axis =>
                axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
        SetMultilineText(horizontalAxis.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!, "Expected", "return", "percent");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Risk" + Environment.NewLine + "Return" +
            Environment.NewLine + "Portfolio", snapshot.Title);
        Assert.Equal("Expected" + Environment.NewLine + "return" +
            Environment.NewLine + "percent", snapshot.Layout.CategoryAxisTitle);
    }

    [Fact]
    public void BubbleChart_RejectsFormulaBackedSeriesNameWithoutCache() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        C.SeriesText seriesText = presentation.Slides[0].SlidePart.ChartParts
            .Single().ChartSpace!.Descendants<C.BubbleChartSeries>().Single()
            .GetFirstChild<C.SeriesText>()!;
        Assert.NotNull(seriesText.GetFirstChild<C.StringReference>());
        seriesText.GetFirstChild<C.StringReference>()!
            .RemoveAllChildren<C.StringCache>();

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void BubbleChart_UpdateRejectsMissingEmbeddedWorkbookBeforeMutation() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        ChartPart chartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
        EmbeddedPackagePart embedded = Assert.Single(
            chartPart.GetPartsOfType<EmbeddedPackagePart>());
        chartPart.DeletePart(embedded);
        string originalChartXml = chartPart.ChartSpace!.OuterXml;

        Assert.Throws<NotSupportedException>(() => chart.UpdateData(CreateData(3D, 5D, 16D)));
        Assert.Equal(originalChartXml, chartPart.ChartSpace!.OuterXml);
    }

    private static PowerPointPresentation CreatePresentation(out PowerPointChart chart) {
        PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        chart = presentation.AddSlide().AddChart(
            OfficeChartKind.Bubble, CreateData(1D, 2D, 4D));
        return presentation;
    }

    private static OfficeChartData CreateData(double x, double y, double size) =>
        new(new[] { x.ToString(System.Globalization.CultureInfo.InvariantCulture) },
            new[] {
                OfficeChartSeries.CreateBubble(
                    "Portfolio", new[] { x }, new[] { y }, new[] { size })
            });

    private static void SetMultilineText(C.ChartText chartText,
        string firstLine, string secondLine, string secondParagraph) {
        C.RichText richText = chartText.GetFirstChild<C.RichText>()!;
        richText.RemoveAllChildren<A.Paragraph>();
        richText.Append(
            new A.Paragraph(
                new A.Run(new A.Text { Text = firstLine }),
                new A.Break(),
                new A.Run(new A.Text { Text = secondLine })),
            new A.Paragraph(
                new A.Run(new A.Text { Text = secondParagraph })));
    }
}
