using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartRadialOutsideLabelsTests {
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
        Assert.Equal(showLeaderLines ? 3 : 0,
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
}
