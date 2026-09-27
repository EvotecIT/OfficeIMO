using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartSecondaryScaleTests {
    [Fact]
    public void SecondaryScaleAndFormatRemainIndependentOfPrimaryAxis() {
        var primaryLayout = new OfficeChartLayout(showLegend: false, verticalAxisMinimum: 0,
            verticalAxisMaximum: 200, verticalAxisNumberFormat: "0.0");
        var layout = primaryLayout.WithSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0,
            maximum: 4, majorUnit: 1, numberFormat: "0%"));
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Primary", new[] { 100d, 200d }),
            new OfficeChartSeries("Secondary", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("Scale", null,
            OfficeChartKind.ColumnClustered, data, 360, 240, layout: layout));
        var labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToArray();
        Assert.Contains("200.0", labels);
        Assert.Contains("400%", labels);
        Assert.DoesNotContain("20000%", labels);
        Assert.Null(primaryLayout.SecondaryValueAxis);
        Assert.Null(OfficeChartLayout.Default.SecondaryValueAxis);
    }
}
