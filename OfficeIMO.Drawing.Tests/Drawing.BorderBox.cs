using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
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

        List<OfficeDrawingShape> lines = EnumerateBorderBoxShapes(drawing)
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

    [Fact]
    public void OfficeDrawingBorderBoxKeepsEdgesOnBoxBounds() {
        var drawing = new OfficeDrawing(40, 30);
        drawing.AddBorderBox(
            0D,
            0D,
            40D,
            30D,
            null,
            new OfficeBorderBox(
                right: new OfficeBorderSide(OfficeColor.FromRgb(37, 99, 235), 1D),
                bottom: new OfficeBorderSide(OfficeColor.FromRgb(22, 163, 74), 1D)));

        List<OfficeDrawingShape> lines = EnumerateBorderBoxShapes(drawing).ToList();

        Assert.Contains(lines, line => line.X == 40D && line.Shape.Width == 0D);
        Assert.Contains(lines, line => line.Y == 30D && line.Shape.Height == 0D);
    }

    [Fact]
    public void OfficeDrawingBorderBoxClipsDoubleEdgesAtDrawingBounds() {
        var drawing = new OfficeDrawing(40, 30);

        drawing.AddBorderBox(
            0D,
            0D,
            40D,
            30D,
            null,
            new OfficeBorderBox(
                left: new OfficeBorderSide(OfficeColor.FromRgb(220, 38, 38), 1D, lineKind: OfficeBorderLineKind.Double, doubleLineSeparation: 8D),
                top: new OfficeBorderSide(OfficeColor.FromRgb(37, 99, 235), 1D, lineKind: OfficeBorderLineKind.Double, doubleLineSeparation: 8D)));

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.White, image.GetPixel(0, 15));
        Assert.Equal(OfficeColor.White, image.GetPixel(15, 0));
        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(220, 38, 38), 4, 15, 128));
        Assert.True(ContainsColorNear(image, OfficeColor.FromRgb(37, 99, 235), 15, 4, 128));
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("viewBox=\"0 0 40 30\"", svg, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 0, 0, 8, 4)]
    [InlineData(10, 12, 0, 8, 4)]
    [InlineData(10, 12, 3, 8, 7)]
    [InlineData(10, 12, 3, 0, 6)]
    [InlineData(10, 12, 8, 8, 12)]
    public void OfficeDrawingBorderBoxCompositionPreservesTheChildViewportThroughSvgAndRaster(int offsetX, int offsetY, int inset, double separation, int innerStroke) {
        var red = new OfficeBorderSide(OfficeColor.Red, 2D, lineKind: OfficeBorderLineKind.Double, doubleLineSeparation: separation);
        var blue = new OfficeBorderSide(OfficeColor.Blue, 2D, lineKind: OfficeBorderLineKind.Double, doubleLineSeparation: separation);
        var child = new OfficeDrawing(40D, 30D).AddBorderBox(
            inset, inset, 40D - (2D * inset), 30D - (2D * inset), null, new OfficeBorderBox(top: blue, right: red, bottom: blue, left: red));
        var drawing = new OfficeDrawing(80D, 70D).AddDrawing(child.Clone(), offsetX, offsetY);

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.Red, image.GetPixel(offsetX + innerStroke, offsetY + 15));
        Assert.Equal(OfficeColor.Red, image.GetPixel(offsetX + 40 - innerStroke, offsetY + 15));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(offsetX + 15, offsetY + innerStroke));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(offsetX + 15, offsetY + 30 - innerStroke));
        Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 44, offsetY + 15));
        Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 15, offsetY + 34));
        Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 40, offsetY + 15));
        Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 15, offsetY + 30));
        if (offsetX > 0) {
            Assert.Equal(OfficeColor.White, image.GetPixel(offsetX - 4, offsetY + 15));
            Assert.Equal(OfficeColor.White, image.GetPixel(offsetX - 1, offsetY + 15));
        }
        if (offsetY > 0) {
            Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 15, offsetY - 4));
            Assert.Equal(OfficeColor.White, image.GetPixel(offsetX + 15, offsetY - 1));
        }

        XDocument svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        XNamespace ns = svg.Root!.Name.Namespace;
        if (inset == 8) {
            Assert.Equal(8, svg.Descendants(ns + "line").Count());
            return;
        }

        XElement clip = Assert.Single(svg.Descendants(ns + "clipPath"));
        XElement viewport = Assert.Single(clip.Elements(ns + "rect"));
        Assert.Equal("0", (string?)viewport.Attribute("x"));
        Assert.Equal("0", (string?)viewport.Attribute("y"));
        Assert.Equal("40", (string?)viewport.Attribute("width"));
        Assert.Equal("30", (string?)viewport.Attribute("height"));
        XElement group = Assert.Single(svg.Descendants(ns + "g"), element =>
            (string?)element.Attribute("clip-path") == "url(#" + (string?)clip.Attribute("id") + ")");
        Assert.Equal($"translate({offsetX} {offsetY})", (string?)group.Attribute("transform"));
    }

    [Fact]
    public void OfficeDrawingBorderBoxRebasesTransformsForEachSide() {
        var drawing = new OfficeDrawing(120, 80);
        OfficeTransform transform = OfficeTransform.RotateDegrees(15D, 40D, 20D);

        drawing.AddBorderBox(
            10D,
            12D,
            80D,
            40D,
            null,
            new OfficeBorderBox(
                top: new OfficeBorderSide(OfficeColor.FromRgb(37, 99, 235), 1D),
                right: new OfficeBorderSide(OfficeColor.FromRgb(22, 163, 74), 1D)),
            transform);

        List<OfficeDrawingShape> lines = EnumerateBorderBoxShapes(drawing).ToList();
        OfficeDrawingShape top = Assert.Single(lines, line => line.X == 10D && line.Y == 12D);
        OfficeDrawingShape right = Assert.Single(lines, line => line.X == 90D);

        Assert.Equal(transform, top.Shape.Transform);
        Assert.NotEqual(transform.OffsetX, right.Shape.Transform!.Value.OffsetX);
        Assert.NotEqual(transform.OffsetY, right.Shape.Transform!.Value.OffsetY);
    }

    private static IEnumerable<OfficeDrawingShape> EnumerateBorderBoxShapes(OfficeDrawing drawing) => drawing.Shapes.Concat(
        drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => EnumerateBorderBoxShapes(group.Drawing)));

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
