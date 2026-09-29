using System.Linq;
using System;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartRadialOutsideLabelsTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void OutsideLabelsRejectNamesThatCannotFitWithoutTruncation(OfficeChartKind kind) {
        using var presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A category name that cannot fit the outside gutter", "B" },
                new[] { new OfficeChartSeries("Status", new[] { 3d, 4d }) }));
        chart.SetDataLabels(showValue: false, showCategoryName: true)
            .SetDataLabelPosition(OfficeChartDataLabelPosition.OutsideEnd);

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Throws<NotSupportedException>(() => OfficeChartDrawingRenderer.Render(snapshot));
    }

    [Fact]
    public void AuthoredEmptyLeaderLinesContainerProjectsOutsideLabels() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        PowerPointChart chart = slide.AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d })
            }));
        chart.SetDataLabels(showValue: false, showCategoryName: true)
            .SetDataLabelPosition(OfficeChartDataLabelPosition.OutsideEnd)
            .SetDataLabelLeaderLines(true);
        C.DataLabels labels = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.PieChart>()
            .Single().GetFirstChild<C.DataLabels>()!;
        Assert.Empty(labels.GetFirstChild<C.LeaderLines>()!.ChildElements);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.True(snapshot.Layout.ShowDataLabelLeaderLines);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeOutsideLabelsRetainLeaderLineSetting(bool showLeaderLines) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
            }));
        C.DoughnutChart native = slide.SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<C.DoughnutChart>().Single();
        var labels = new C.DataLabels();
        labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
        labels.AddChild(new C.ShowCategoryName { Val = true }, true);
        labels.AddChild(new C.ShowLeaderLines { Val = showLeaderLines }, true);
        native.AddChild(labels, true);

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartDataLabelPosition.OutsideEnd, snapshot.Layout.DataLabelPosition);
        Assert.Equal(showLeaderLines, snapshot.Layout.ShowDataLabelLeaderLines);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(showLeaderLines ? 6 : 0,
            drawing.Shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Line));
        foreach (string category in new[] { "A", "B", "C" })
            Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.Contains(category));
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void NativeOutsideLabelsUseChartDefaultFontSize() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d })
            }));
        C.ChartSpace space = slide.SlidePart.ChartParts.Single().ChartSpace!;
        space.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties { FontSize = 1800 }))), true);
        C.DoughnutChart native = space.Descendants<C.DoughnutChart>().Single();
        var labels = new C.DataLabels();
        labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
        labels.AddChild(new C.ShowCategoryName { Val = true }, true);
        native.AddChild(labels, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(18D, snapshot.Layout.DataLabelFontSize);
    }

    [Fact]
    public void NativeDataLabelsPreserveExplicitTypeface() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d })
            }));
        chart.SetDataLabels(showValue: true);
        ChartPart chartPart = slide.SlidePart.ChartParts.Single();
        foreach (ChartStylePart stylePart in chartPart.GetPartsOfType<ChartStylePart>().ToArray())
            chartPart.DeletePart(stylePart);
        C.DataLabels labels = chartPart.ChartSpace!
            .Descendants<C.BarChart>().Single().GetFirstChild<C.DataLabels>()!;
        labels.GetFirstChild<C.TextProperties>()?.Remove();
        labels.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.LatinFont { Typeface = "Georgia" })))), true);
        Assert.Contains("Georgia", labels.OuterXml);
        Assert.Equal("Georgia", labels.GetFirstChild<C.TextProperties>()?
            .Descendants<A.LatinFont>().Single().Typeface?.Value);

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Georgia", snapshot.Layout.DataLabelFontFamily);
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Elements.OfType<OfficeDrawingText>(),
            text => text.Font.FamilyName == "Georgia");
    }

    [Fact]
    public void DataUpdatePreservesUnprojectedRadialPointLabels() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 3d, 4d })
        });
        PowerPointChart chart = slide.AddChart(OfficeChartKind.Pie, data);
        C.PieChart native = slide.SlidePart.ChartParts.Single().ChartSpace!
            .Descendants<C.PieChart>().Single();
        var labels = new C.DataLabels();
        labels.AddChild(new C.ShowValue { Val = true }, true);
        labels.AddChild(new C.DataLabel(new C.Index { Val = 0U },
            new C.ShowValue { Val = false }), true);
        native.AddChild(labels, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        chart.UpdateData(new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 5d, 6d })
        }));
        Assert.NotNull(native.GetFirstChild<C.DataLabels>()?.GetFirstChild<C.DataLabel>());
    }
}
