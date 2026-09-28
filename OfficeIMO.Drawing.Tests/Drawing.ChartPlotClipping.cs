using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartPlotClippingTests {
    [Fact]
    public void InRangePointLabelCanExtendAboveClippedPlot() {
        var data = new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Series", new[] { 0d, 8d, 10d })
        });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelValues: true, dataLabelNumberFormat: "0.00",
            dataLabelPosition: OfficeChartDataLabelPosition.Top,
            verticalAxisMinimum: 2, verticalAxisMaximum: 8);
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Line, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingText label = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>(),
            text => text.Text == "8.00");
        Assert.True(label.Y < group.Y, "The in-range point label should be allowed above the plot.");
        Assert.DoesNotContain(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "0.00" || text.Text == "10.00");
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.Scatter)]
    public void ExplicitAxisBoundsClipSeriesPaintWithoutClippingLabels(OfficeChartKind kind) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var series = kind == OfficeChartKind.Scatter
            ? new OfficeChartSeries("Series", new[] { 5d, 5d, 5d }, new[] { -2d, 5d, 12d }, ink)
            : new OfficeChartSeries("Series", new[] { 0d, 5d, 10d }, null, ink);
        var data = new OfficeChartData(new[] { "A", "B", "C" }, new[] { series });
        var layout = new OfficeChartLayout(showLegend: false, showDataLabels: true,
            showDataLabelValues: true, dataLabelNumberFormat: "0.00",
            verticalAxisMinimum: kind == OfficeChartKind.Scatter ? 0 : 2,
            verticalAxisMaximum: kind == OfficeChartKind.Scatter ? 10 : 8,
            horizontalAxisMinimum: kind == OfficeChartKind.Scatter ? 0 : null,
            horizontalAxisMaximum: kind == OfficeChartKind.Scatter ? 10 : null);
        var style = new OfficeChartStyle(showBackground: false, palette: new[] { ink },
            showGridLines: false, showBorder: false);
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 360, 240, style, layout));

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "5.00");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        int inside = 0;
        int outside = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++) {
                if (raster.GetPixel(x, y) != ink) continue;
                if (x >= group.X && x < group.X + group.ClipPath.Width &&
                    y >= group.Y && y < group.Y + group.ClipPath.Height) inside++;
                else outside++;
            }
        Assert.True(inside > 0, "Expected visible series paint inside the plot.");
        Assert.Equal(0, outside);
    }
}
