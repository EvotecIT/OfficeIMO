using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingChartPlotClippingTests {
    [Fact]
    public void AutomaticScatterRangeMapsOppositeFiniteExtremes() {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", new[] { -double.MaxValue, double.MaxValue },
                new[] { -double.MaxValue, double.MaxValue }, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false);

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, data, 360, 240, layout: layout));

        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Line &&
            shape.Shape.StrokeColor == ink && shape.Shape.Width > 100D && shape.Shape.Height > 100D);
        Assert.DoesNotContain(drawing.Elements.OfType<OfficeDrawingText>(), text =>
            text.Text.Contains("NaN", StringComparison.Ordinal) ||
            text.Text.Contains("Infinity", StringComparison.Ordinal));
    }

    [Fact]
    public void ClippedAreaDoesNotOutlineTheArtificialUpperEdge() {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", new[] { 10d, 10d }, null, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            verticalAxisMinimum: 0, verticalAxisMaximum: 5);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Area, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>());
        Assert.Contains(group.Drawing.Shapes, shape => shape.Shape.FillColor == ink);
        Assert.DoesNotContain(group.Drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Polygon && shape.Shape.StrokeColor == ink);
        Assert.DoesNotContain(group.Drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink &&
            Math.Abs(shape.Y - group.Y) < .01D && shape.Shape.Width > group.ClipPath.Width * .9D);
        Assert.Contains(group.Drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink &&
            shape.Shape.Width > group.ClipPath.Width * .9D &&
            shape.Y > group.Y + group.ClipPath.Height * .9D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OppositeFiniteExtremeScatterPointsCrossTheEntirePlot(bool diagonal) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", diagonal
                    ? new[] { -double.MaxValue, double.MaxValue }
                    : new[] { .5d, .5d },
                new[] { -double.MaxValue, double.MaxValue }, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 1,
            verticalAxisMinimum: 0, verticalAxisMaximum: 1);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(),
            item => item.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == ink));
        OfficeDrawingShape line = Assert.Single(group.Drawing.Shapes,
            shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink);
        Assert.InRange(Math.Abs(line.X - group.X), 0D, .01D);
        Assert.True(line.Shape.Width > group.ClipPath.Width * .99d);
        if (diagonal) Assert.True(line.Shape.Height > group.ClipPath.Height * .99d);
        else {
            Assert.InRange(Math.Abs(line.Y - (group.Y + group.ClipPath.Height * .5d)), 0D, 1D);
            Assert.InRange(line.Shape.Height, 0D, 1D);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FiniteExtremeScatterKeepsConnectedLineFromUpperBoundary(bool reverse) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        double[] x = reverse ? new[] { .5d, double.MaxValue } : new[] { double.MaxValue, .5d };
        double[] y = reverse ? new[] { .5d, double.MaxValue } : new[] { double.MaxValue, .5d };
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", y, x, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 1,
            verticalAxisMinimum: 0, verticalAxisMaximum: 1);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(),
            item => item.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == ink));
        OfficeDrawingShape line = Assert.Single(group.Drawing.Shapes,
            shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink);
        Assert.True(line.Shape.Width > group.ClipPath.Width * .45d);
        Assert.True(line.Shape.Height > group.ClipPath.Height * .45d);
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.Scatter)]
    public void FiniteExtremeOutlierClipsBeforePixelConversion(OfficeChartKind kind) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var series = kind == OfficeChartKind.Scatter
            ? new OfficeChartSeries("Series", new[] { double.MaxValue, .5d },
                new[] { double.MaxValue, .5d }, ink)
            : new OfficeChartSeries("Series", new[] { double.MaxValue, .5d }, null, ink);
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 1,
            verticalAxisMinimum: 0, verticalAxisMaximum: 1);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 360, 240, layout: layout));
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingGroup>(), group =>
            group.Drawing.Shapes.Any(shape => shape.Shape.StrokeColor == ink || shape.Shape.FillColor == ink));
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void OneSidedBoundOutsideDataStillClipsSeries(OfficeChartKind kind) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var series = kind == OfficeChartKind.Scatter
            ? new OfficeChartSeries("Series", new[] { 10d, 20d }, new[] { 0d, 1d }, ink)
            : new OfficeChartSeries("Series", new[] { 10d, 20d }, null, ink);
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { series });
        var layout = new OfficeChartLayout(showLegend: false, verticalAxisMaximum: 5);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 360, 240, layout: layout));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                Assert.NotEqual(ink, raster.GetPixel(x, y));
    }

    [Fact]
    public void ExtremeScatterOutlierKeepsTrueLineSlopeAtPlotBoundary() {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", new[] { -1e9, .5d }, new[] { -1e12, .5d }, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            horizontalAxisMinimum: 0, horizontalAxisMaximum: 1,
            verticalAxisMinimum: 0, verticalAxisMaximum: 1);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.Scatter, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(),
            item => item.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == ink));
        OfficeDrawingShape line = Assert.Single(group.Drawing.Shapes,
            shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink);
        Assert.InRange(Math.Abs(line.X - group.X), 0D, .01D);
        Assert.InRange(Math.Abs(line.Y - (group.Y + group.ClipPath.Height * .5d)), 0D, 1D);
        Assert.True(line.Shape.Width > group.ClipPath.Width * .4d);
        Assert.InRange(line.Shape.Height, 0D, 1D);
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    public void ExtremeCategoryOutliersKeepTrueCrossingNearSecondCategory(OfficeChartKind kind) {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Series", new[] { -1e12, 1e9 }, null, ink)
        });
        var layout = new OfficeChartLayout(showLegend: false,
            verticalAxisMinimum: 0, verticalAxisMaximum: 1);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            kind, data, 360, 240, layout: layout));
        OfficeDrawingGroup group = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(),
            item => item.Drawing.Shapes.Any(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == ink));
        OfficeDrawingShape line = Assert.Single(group.Drawing.Shapes,
            shape => shape.Shape.Kind == OfficeShapeKind.Line && shape.Shape.StrokeColor == ink &&
                shape.Shape.StrokeWidth > 0.5D);
        Assert.True(line.X > group.X + group.ClipPath.Width * .9d);
        Assert.True(line.Shape.Width < group.ClipPath.Width * .1d);
    }

    [Fact]
    public void HorizontalBarBoundsClipMixedPrimaryLine() {
        OfficeColor ink = OfficeColor.Parse("#D900AA");
        var data = new OfficeChartData(new[] { "A", "B" }, new OfficeChartSeries[] {
            new("Bars", new[] { 1d, 2d }),
            new("Trend", new[] { 10d, 20d }, null, ink, null, true,
                renderKind: OfficeChartKind.Line)
        });
        var layout = new OfficeChartLayout(showLegend: false, horizontalAxisMaximum: 5);
        var drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot("", null,
            OfficeChartKind.BarClustered, data, 360, 240, layout: layout));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        for (int y = 0; y < raster.Height; y++)
            for (int x = 0; x < raster.Width; x++)
                Assert.NotEqual(ink, raster.GetPixel(x, y));
    }

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
