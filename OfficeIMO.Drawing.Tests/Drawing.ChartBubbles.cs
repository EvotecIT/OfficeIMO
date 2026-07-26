using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingChartBubbleTests {
    [Fact]
    public void OfficeChartSnapshot_RetainsBinaryCompatibleConstructor() {
        Assert.NotNull(typeof(OfficeChartSnapshot).GetConstructor(new[] {
            typeof(string),
            typeof(string),
            typeof(OfficeChartKind),
            typeof(OfficeChartData),
            typeof(double),
            typeof(double),
            typeof(OfficeChartStyle),
            typeof(OfficeChartLayout)
        }));
    }

    [Fact]
    public void OfficeChartDrawingRenderer_HonorsBubbleScaleAndSizeMode() {
        OfficeDrawing area = RenderBubbles(100D, OfficeChartBubbleSizeMode.Area);
        OfficeDrawing width = RenderBubbles(100D, OfficeChartBubbleSizeMode.Width);
        OfficeDrawing scaled = RenderBubbles(200D, OfficeChartBubbleSizeMode.Area);

        double[] areaDiameters = GetBubbleDiameters(area);
        double[] widthDiameters = GetBubbleDiameters(width);
        double[] scaledDiameters = GetBubbleDiameters(scaled);

        Assert.Equal(2, areaDiameters.Length);
        Assert.Equal(2, widthDiameters.Length);
        Assert.Equal(2, scaledDiameters.Length);
        Assert.Equal(areaDiameters[1] * 0.5D, areaDiameters[0], precision: 6);
        Assert.Equal(widthDiameters[1] * 0.25D, widthDiameters[0], precision: 6);
        Assert.Equal(areaDiameters[1], widthDiameters[1], precision: 6);
        Assert.Equal(areaDiameters[1] * 2D, scaledDiameters[1], precision: 6);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_InsetsBubbleExtremaInsidePlotAxes() {
        OfficeDrawing drawing = RenderBubbles(200D, OfficeChartBubbleSizeMode.Area);
        OfficeDrawingShape horizontalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width > 100D && shape.Shape.Height == 0D)
            .OrderByDescending(shape => shape.Shape.Width)
            .First();
        OfficeDrawingShape verticalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width == 0D && shape.Shape.Height > 50D)
            .OrderByDescending(shape => shape.Shape.Height)
            .First();

        foreach (OfficeDrawingShape bubble in GetBubbles(drawing)) {
            Assert.True(bubble.X >= horizontalAxis.X - 0.001D);
            Assert.True(bubble.X + bubble.Shape.Width <=
                        horizontalAxis.X + horizontalAxis.Shape.Width + 0.001D);
            Assert.True(bubble.Y >= verticalAxis.Y - 0.001D);
            Assert.True(bubble.Y + bubble.Shape.Height <=
                        verticalAxis.Y + verticalAxis.Shape.Height + 0.001D);
        }

        OfficeDrawingShape[] bubbles = GetBubbles(drawing)
            .OrderBy(shape => shape.X + shape.Shape.Width / 2D)
            .ToArray();
        double[] horizontalTickPositions = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width == 0D &&
                            shape.Shape.Height > 0D &&
                            shape.Shape.Height <= 4.001D)
            .Select(shape => shape.X)
            .Distinct()
            .OrderBy(value => value)
            .ToArray();
        double[] verticalTickPositions = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Height == 0D &&
                            shape.Shape.Width > 0D &&
                            shape.Shape.Width <= 4.001D)
            .Select(shape => shape.Y)
            .Distinct()
            .OrderBy(value => value)
            .ToArray();

        Assert.Equal(5, horizontalTickPositions.Length);
        Assert.Equal(5, verticalTickPositions.Length);
        Assert.Equal(horizontalTickPositions[0],
            bubbles[0].X + bubbles[0].Shape.Width / 2D, precision: 6);
        Assert.Equal(horizontalTickPositions[4],
            bubbles[1].X + bubbles[1].Shape.Width / 2D, precision: 6);
        Assert.Equal(verticalTickPositions[0],
            bubbles[1].Y + bubbles[1].Shape.Height / 2D, precision: 6);
        Assert.Equal(verticalTickPositions[4],
            bubbles[0].Y + bubbles[0].Shape.Height / 2D, precision: 6);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_InsetsBubbleOutlinesInsidePlotAxes() {
        OfficeColor color = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble(
                "Outlined",
                new[] { 1D, 2D },
                new[] { 1D, 2D },
                new[] { 25D, 100D },
                color,
                markerOutlineColor: OfficeColor.Parse("#112233"),
                markerOutlineWidth: 12D)
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot(
                "Outlined bubbles",
                null,
                OfficeChartKind.Bubble,
                data,
                widthPoints: 420D,
                heightPoints: 260D,
                style: null,
                layout: new OfficeChartLayout(showLegend: false),
                bubbleScalePercent: 200D,
                bubbleSizeMode: OfficeChartBubbleSizeMode.Area));
        OfficeDrawingShape horizontalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width > 100D &&
                            shape.Shape.Height == 0D)
            .OrderByDescending(shape => shape.Shape.Width)
            .First();
        OfficeDrawingShape verticalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width == 0D &&
                            shape.Shape.Height > 50D)
            .OrderByDescending(shape => shape.Shape.Height)
            .First();

        Assert.All(GetBubbles(drawing), bubble => {
            double halfStroke = bubble.Shape.StrokeWidth / 2D;
            Assert.True(bubble.X - halfStroke >=
                        horizontalAxis.X - 0.001D,
                $"Left edge {bubble.X - halfStroke} is before axis {horizontalAxis.X}.");
            Assert.True(bubble.X + bubble.Shape.Width + halfStroke <=
                        horizontalAxis.X + horizontalAxis.Shape.Width + 0.001D,
                $"Right edge {bubble.X + bubble.Shape.Width + halfStroke} exceeds axis {horizontalAxis.X + horizontalAxis.Shape.Width}.");
            Assert.True(bubble.Y - halfStroke >=
                        verticalAxis.Y - 0.001D,
                $"Top edge {bubble.Y - halfStroke} is above axis {verticalAxis.Y}.");
            Assert.True(bubble.Y + bubble.Shape.Height + halfStroke <=
                        verticalAxis.Y + verticalAxis.Shape.Height + 0.001D,
                $"Bottom edge {bubble.Y + bubble.Shape.Height + halfStroke} exceeds axis {verticalAxis.Y + verticalAxis.Shape.Height}.");
        });
    }

    [Fact]
    public void OfficeChartDrawingRenderer_InsetsBubblesInMixedScatterCharts() {
        OfficeColor bubbleColor = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "1", "2" }, new OfficeChartSeries[] {
            new OfficeChartSeries(
                "Scatter",
                new[] { 1D, 2D },
                new[] { 1D, 2D },
                OfficeColor.Parse("#264653"),
                pointColors: null,
                showMarkers: true,
                renderKind: OfficeChartKind.Scatter),
            OfficeChartSeries.CreateBubble(
                "Bubbles",
                new[] { 1D, 2D },
                new[] { 1D, 2D },
                new[] { 25D, 100D },
                bubbleColor)
        });
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot(
                "Mixed scatter",
                null,
                OfficeChartKind.Scatter,
                data,
                widthPoints: 420D,
                heightPoints: 260D,
                style: null,
                layout: new OfficeChartLayout(showLegend: false),
                bubbleScalePercent: 200D,
                bubbleSizeMode: OfficeChartBubbleSizeMode.Area));
        OfficeDrawingShape horizontalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width > 100D && shape.Shape.Height == 0D)
            .OrderByDescending(shape => shape.Shape.Width)
            .First();
        OfficeDrawingShape verticalAxis = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line &&
                            shape.Shape.Width == 0D && shape.Shape.Height > 50D)
            .OrderByDescending(shape => shape.Shape.Height)
            .First();

        OfficeDrawingShape[] bubbles = drawing.Shapes.Where(shape =>
                shape.Shape.Kind == OfficeShapeKind.Ellipse &&
                shape.Shape.FillColor == bubbleColor)
            .ToArray();
        Assert.Equal(2, bubbles.Length);
        Assert.All(bubbles, bubble => {
            Assert.True(bubble.X >= horizontalAxis.X - 0.001D);
            Assert.True(bubble.X + bubble.Shape.Width <=
                        horizontalAxis.X + horizontalAxis.Shape.Width + 0.001D);
            Assert.True(bubble.Y >= verticalAxis.Y - 0.001D);
            Assert.True(bubble.Y + bubble.Shape.Height <=
                        verticalAxis.Y + verticalAxis.Shape.Height + 0.001D);
        });
    }

    [Fact]
    public void OfficeChartDrawingRenderer_RendersBubblesWhenScatterMarkersAreHidden() {
        OfficeColor color = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble("Portfolio",
                new[] { 1D, 2D },
                new[] { 1D, 2D },
                new[] { 25D, 100D },
                color)
        });

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot(
                "Bubbles",
                null,
                OfficeChartKind.Bubble,
                data,
                widthPoints: 420D,
                heightPoints: 260D,
                layout: new OfficeChartLayout(showLegend: false, showMarkers: false)));

        Assert.Equal(2, GetBubbles(drawing).Length);
    }

    [Fact]
    public void OfficeChartDrawingRenderer_PreservesDisabledBubbleOutline() {
        OfficeColor color = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "1" }, new[] {
            OfficeChartSeries.CreateBubble("Portfolio",
                new[] { 1D },
                new[] { 2D },
                new[] { 25D },
                color,
                showMarkerOutline: false)
        });

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(
            new OfficeChartSnapshot(
                "Bubbles",
                null,
                OfficeChartKind.Bubble,
                data,
                widthPoints: 420D,
                heightPoints: 260D));
        OfficeDrawingShape bubble = Assert.Single(GetBubbles(drawing));

        Assert.False(Assert.Single(data.Series).ShowMarkerOutline);
        Assert.Null(bubble.Shape.StrokeColor);
        Assert.Equal(0D, bubble.Shape.StrokeWidth);
    }

    [Fact]
    public void OfficeChartSnapshot_RejectsBubbleSeriesWithoutSizes() {
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries(
                "Incomplete",
                new[] { 2D, 4D },
                new[] { 1D, 2D })
        });

        Assert.Throws<System.ArgumentException>(() => new OfficeChartSnapshot(
            "Malformed",
            null,
            OfficeChartKind.Bubble,
            data,
            widthPoints: 420D,
            heightPoints: 260D));
    }

    [Fact]
    public void OfficeChartSnapshot_RejectsMixedBubbleSeriesWithoutSizes() {
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            new OfficeChartSeries(
                "Incomplete bubble",
                new[] { 2D, 4D },
                new[] { 1D, 2D },
                color: null,
                pointColors: null,
                showMarkers: true,
                renderKind: OfficeChartKind.Bubble)
        });

        Assert.Throws<System.ArgumentException>(() => new OfficeChartSnapshot(
            "Malformed mixed scatter",
            null,
            OfficeChartKind.Scatter,
            data,
            widthPoints: 420D,
            heightPoints: 260D));
    }

    [Fact]
    public void OfficeChartSnapshot_RejectsBubbleSeriesOnCategoricalAxes() {
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble(
                "Misplaced bubble",
                new[] { 1D, 2D },
                new[] { 3D, 4D },
                new[] { 5D, 6D })
        });

        Assert.Throws<System.ArgumentException>(() => new OfficeChartSnapshot(
            "Categorical",
            null,
            OfficeChartKind.ColumnClustered,
            data,
            widthPoints: 420D,
            heightPoints: 260D));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void OfficeChartSeries_RejectsNonFiniteBubbleOutlineWidths(
        double outlineWidth) {
        Assert.Throws<System.ArgumentOutOfRangeException>(() =>
            OfficeChartSeries.CreateBubble(
                "Invalid outline",
                new[] { 1D },
                new[] { 2D },
                new[] { 4D },
                markerOutlineWidth: outlineWidth));
    }

    private static OfficeDrawing RenderBubbles(double scale,
        OfficeChartBubbleSizeMode sizeMode) {
        OfficeColor color = OfficeColor.Parse("#2A9D8F");
        var data = new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble("Portfolio",
                new[] { 1D, 2D },
                new[] { 1D, 2D },
                new[] { 25D, 100D },
                color)
        });
        return OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Bubbles",
            null,
            OfficeChartKind.Bubble,
            data,
            widthPoints: 420D,
            heightPoints: 260D,
            style: null,
            layout: new OfficeChartLayout(
                showLegend: false,
                horizontalAxisMajorTickMark: OfficeChartAxisTickMark.Cross,
                verticalAxisMajorTickMark: OfficeChartAxisTickMark.Cross),
            bubbleScalePercent: scale,
            bubbleSizeMode: sizeMode));
    }

    private static OfficeDrawingShape[] GetBubbles(OfficeDrawing drawing) =>
        drawing.Shapes.Where(shape =>
                shape.Shape.Kind == OfficeShapeKind.Ellipse &&
                shape.Shape.FillColor == OfficeColor.Parse("#2A9D8F"))
            .OrderBy(shape => shape.Shape.Width)
            .ToArray();

    private static double[] GetBubbleDiameters(OfficeDrawing drawing) =>
        GetBubbles(drawing).Select(shape => shape.Shape.Width).ToArray();
}
