using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartPlotClippingTests {
    [Fact]
    public void MixedBubblePaddingDoesNotExposeOffscaleScatterLine() {
        OfficeColor lineInk = OfficeColor.Parse("#D900AA");
        OfficeColor bubbleInk = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "A", "B" }, new OfficeChartSeries[] {
            new("Trend", new[] { 5d, 5d }, new[] { -2d, 5d }, lineInk),
            OfficeChartSeries.CreateBubble("Bubbles", new[] { 0d, 10d },
                new[] { 0d, 10d }, new[] { 100d, 100d }, bubbleInk)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 10,
            verticalAxisMinimum: 0, verticalAxisMaximum: 10);
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, data, 420, 260, layout: layout));

        OfficeDrawingGroup[] groups = drawing.Elements.OfType<OfficeDrawingGroup>().ToArray();
        OfficeDrawingGroup lineGroup = Assert.Single(groups,
            group => group.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == lineInk));
        OfficeDrawingGroup bubbleGroup = Assert.Single(groups,
            group => group.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Ellipse &&
                shape.Shape.FillColor == bubbleInk));
        Assert.True(bubbleGroup.X < lineGroup.X, "Bubbles reserve an outer margin around the numeric plot.");

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        int marginInk = 0;
        int plottedInk = 0;
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++) {
                if (raster.GetPixel(x, y) != lineInk) continue;
                if (x < lineGroup.X) marginInk++;
                else plottedInk++;
            }
        Assert.Equal(0, marginInk);
        Assert.True(plottedInk > 0, "The in-range part of the connected scatter line remains visible.");
    }

    [Fact]
    public void BubbleAtExplicitAxisBoundRetainsPaintInReservedPlotMargin() {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            OfficeChartSeries.CreateBubble("Bubbles", new[] { 0d, 10d },
                new[] { 0d, 10d }, new[] { 100d, 100d }, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 10,
            verticalAxisMinimum: 0, verticalAxisMaximum: 10);
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Bubble, data, 420, 260, layout: layout));

        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        OfficeDrawingShape firstBubble = group.Drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Ellipse && shape.Shape.FillColor == ink)
            .OrderBy(shape => shape.X).First();
        double centerX = firstBubble.X + firstBubble.Shape.Width / 2D;
        Assert.True(firstBubble.X < centerX && firstBubble.X >= group.X);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        int leftInteriorX = (int)Math.Ceiling(firstBubble.X + firstBubble.Shape.Width * 0.25D);
        int centerY = (int)Math.Round(firstBubble.Y + firstBubble.Shape.Height / 2D);
        Assert.Equal(ink, raster.GetPixel(leftInteriorX, centerY));
    }

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

        OfficeDrawingGroup group = drawing.Elements.OfType<OfficeDrawingGroup>().First();
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
