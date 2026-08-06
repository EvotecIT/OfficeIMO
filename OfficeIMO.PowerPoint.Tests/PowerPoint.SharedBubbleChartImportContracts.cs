using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

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

    [Fact]
    public void BubbleChart_HidesLegendEntriesByPlottedOrder() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        var data = new OfficeChartData(new[] { "1" }, new[] {
            OfficeChartSeries.CreateBubble(
                "First", new[] { 1D }, new[] { 2D }, new[] { 4D }),
            OfficeChartSeries.CreateBubble(
                "Second", new[] { 3D }, new[] { 4D }, new[] { 9D })
        });
        PowerPointChart chart = presentation.AddSlide().AddChart(
            OfficeChartKind.Bubble, data);
        C.Chart nativeChart = presentation.Slides[0].SlidePart.ChartParts
            .Single().ChartSpace!.GetFirstChild<C.Chart>()!;
        C.BubbleChartSeries[] series = nativeChart
            .Descendants<C.BubbleChartSeries>().ToArray();
        series[0].Index!.Val = 5U;
        series[0].Order!.Val = 1U;
        series[1].Index!.Val = 9U;
        series[1].Order!.Val = 0U;
        C.Legend legend = nativeChart.GetFirstChild<C.Legend>()!;
        legend.InsertAfter(new C.LegendEntry(
            new C.Index { Val = 0U }, new C.Delete { Val = true }),
            legend.GetFirstChild<C.LegendPosition>());

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Second", snapshot.Data.Series[0].Name);
        Assert.False(snapshot.Data.Series[0].ShowInLegend);
        Assert.True(snapshot.Data.Series[1].ShowInLegend);
    }

    [Fact]
    public void BubbleChart_RejectsInsetSeriesOutlineAlignment() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        A.Outline outline = presentation.Slides[0].SlidePart.ChartParts
            .Single().ChartSpace!.Descendants<C.BubbleChartSeries>().Single()
            .GetFirstChild<C.ChartShapeProperties>()!
            .GetFirstChild<A.Outline>()!;

        outline.Alignment = A.PenAlignmentValues.Insert;
        Assert.False(chart.TryGetOfficeSnapshot(out _));

        outline.Alignment = A.PenAlignmentValues.Center;
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void BubbleChart_RejectsVisibleOnlySnapshotWithHiddenWorkbookRows() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        ChartPart chartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
        EmbeddedPackagePart embedded = Assert.Single(
            chartPart.GetPartsOfType<EmbeddedPackagePart>());
        using (Stream stream = embedded.GetStream(FileMode.Open, FileAccess.ReadWrite))
        using (SpreadsheetDocument workbook = SpreadsheetDocument.Open(stream, true)) {
            S.Row row = workbook.WorkbookPart!.WorksheetParts.Single()
                .Worksheet!.Descendants<S.Row>().Single(item =>
                    item.RowIndex?.Value == 2U);
            row.Hidden = true;
            row.Ancestors<S.Worksheet>().Single().Save();
        }

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void BubbleChart_RecreatedDataLabelsPrecedeBubble3D() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);

        chart.ClearDataLabels().SetDataLabels();

        ChartPart chartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
        C.BubbleChart nativeChart = chartPart.ChartSpace!
            .Descendants<C.BubbleChart>().Single();
        Assert.True(nativeChart.ChildElements.ToList().IndexOf(
            nativeChart.GetFirstChild<C.DataLabels>()!) <
            nativeChart.ChildElements.ToList().IndexOf(
                nativeChart.GetFirstChild<C.Bubble3D>()!));
        Assert.Empty(new OpenXmlValidator().Validate(chartPart));
    }

    [Fact]
    public void BubbleChart_UpdatedPointColorsPrecedePreservedTrendline() {
        using PowerPointPresentation presentation = CreatePresentation(out PowerPointChart chart);
        chart.SetSeriesTrendline(0, OfficeChartTrendlineType.Linear);

        chart.UpdateData(CreateData(3D, 5D, 16D,
            new OfficeColor?[] { OfficeColor.Parse("#445566") }));

        ChartPart chartPart = presentation.Slides[0].SlidePart.ChartParts.Single();
        C.BubbleChartSeries series = chartPart.ChartSpace!
            .Descendants<C.BubbleChartSeries>().Single();
        Assert.True(series.ChildElements.ToList().IndexOf(
            series.GetFirstChild<C.DataPoint>()!) <
            series.ChildElements.ToList().IndexOf(
                series.GetFirstChild<C.Trendline>()!));
        Assert.Empty(new OpenXmlValidator().Validate(chartPart));
    }

    private static PowerPointPresentation CreatePresentation(out PowerPointChart chart) {
        PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        chart = presentation.AddSlide().AddChart(
            OfficeChartKind.Bubble, CreateData(1D, 2D, 4D));
        return presentation;
    }

    private static OfficeChartData CreateData(double x, double y, double size,
        OfficeColor?[]? pointColors = null) =>
        new(new[] { x.ToString(System.Globalization.CultureInfo.InvariantCulture) },
            new[] {
                OfficeChartSeries.CreateBubble(
                    "Portfolio", new[] { x }, new[] { y }, new[] { size },
                    pointColors: pointColors)
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
