using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingChartBubbleTests {
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
        Assert.True(widthDiameters[0] < areaDiameters[0]);
        Assert.Equal(areaDiameters[1], widthDiameters[1], precision: 6);
        Assert.True(scaledDiameters[1] > areaDiameters[1] * 1.9D);
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
            layout: new OfficeChartLayout(showLegend: false),
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
