using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingBorderBoxRendersIndependentEdgesThroughSvgAndRaster() {
        var drawing = new OfficeDrawing(120, 80);
        drawing.AddBorderBox(
            10,
            12,
            80,
            40,
            OfficeColor.FromRgb(250, 250, 250),
            new OfficeBorderBox(
                new OfficeBorderSide(OfficeColor.FromRgb(220, 38, 38), 2D),
                new OfficeBorderSide(OfficeColor.FromRgb(37, 99, 235), 1.5D, OfficeStrokeDashStyle.Dash),
                new OfficeBorderSide(OfficeColor.FromRgb(22, 163, 74), 3D, OfficeStrokeDashStyle.Dot),
                new OfficeBorderSide(OfficeColor.FromRgb(147, 51, 234), 2.5D, OfficeStrokeDashStyle.DashDot)));

        List<OfficeDrawingShape> lines = drawing.Elements
            .OfType<OfficeDrawingShape>()
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
            .ToList();
        Assert.Equal(4, lines.Count);
        Assert.Contains(lines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 2D);
        Assert.Contains(lines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(37, 99, 235) && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
        Assert.Contains(lines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(22, 163, 74) && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dot);
        Assert.Contains(lines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(147, 51, 234) && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.DashDot);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("#DC2626", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#2563EB", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#16A34A", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#9333EA", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("stroke-dasharray", svg, StringComparison.Ordinal);

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(120, image.Width);
        Assert.Equal(80, image.Height);
    }

    [Fact]
    public void OfficeBorderBoxRendererRendersDiagonalEdgesThroughSvgAndRaster() {
        var borders = new OfficeBorderBox(
            diagonalDown: new OfficeBorderSide(OfficeColor.FromRgb(220, 38, 38), 2D, OfficeStrokeDashStyle.DashDot),
            diagonalUp: new OfficeBorderSide(OfficeColor.FromRgb(37, 99, 235), 2D, OfficeStrokeDashStyle.Dot));
        var builder = new StringBuilder();
        OfficeBorderBoxRenderer.AppendSvg(builder, 4D, 6D, 70D, 42D, borders);

        string svg = builder.ToString();
        Assert.Contains("#DC2626", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#2563EB", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("stroke-dasharray", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linecap=\"round\"", svg, StringComparison.Ordinal);

        var image = new OfficeRasterImage(90, 60, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image);
        OfficeBorderBoxRenderer.DrawRaster(canvas, 4D, 6D, 70D, 42D, borders);

        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(220, 38, 38), 25, 18, 18));
        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(37, 99, 235), 25, 36, 18));
    }

    [Fact]
    public void OfficeBorderBoxRendererRendersDoubleEdgesThroughSvgAndRaster() {
        var borders = new OfficeBorderBox(
            left: new OfficeBorderSide(OfficeColor.FromRgb(220, 38, 38), 2D, lineKind: OfficeBorderLineKind.Double, doubleLineSeparation: 8D));
        var builder = new StringBuilder();
        OfficeBorderBoxRenderer.AppendSvg(builder, 20D, 10D, 50D, 36D, borders);

        string svg = builder.ToString();
        Assert.True(CountOccurrences(svg, "#DC2626") >= 2);
        Assert.Contains("x1=\"16\"", svg, StringComparison.Ordinal);
        Assert.Contains("x1=\"24\"", svg, StringComparison.Ordinal);

        var image = new OfficeRasterImage(90, 60, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image);
        OfficeBorderBoxRenderer.DrawRaster(canvas, 20D, 10D, 50D, 36D, borders);

        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(220, 38, 38), 16, 28, 4));
        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(220, 38, 38), 24, 28, 4));
    }

    private static bool ContainsColorNear(OfficeRasterImage image, OfficeColor expected, int centerX, int centerY, int tolerance) {
        int left = Math.Max(0, centerX - 6);
        int right = Math.Min(image.Width - 1, centerX + 6);
        int top = Math.Max(0, centerY - 6);
        int bottom = Math.Min(image.Height - 1, centerY + 6);
        for (int y = top; y <= bottom; y++) {
            for (int x = left; x <= right; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (Math.Abs(pixel.R - expected.R) <= tolerance &&
                    Math.Abs(pixel.G - expected.G) <= tolerance &&
                    Math.Abs(pixel.B - expected.B) <= tolerance &&
                    pixel.A > 160) {
                    return true;
                }
            }
        }

        return false;
    }
}
