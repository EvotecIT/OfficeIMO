using OfficeIMO.Drawing;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Checks the paint retained by a PDF-backed SVG independently of its markup representation.</summary>
internal static class HtmlSvgPaintAssertions {
    internal static OfficeDrawing Read(byte[] bytes) {
        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.NotNull(drawing);
        return drawing;
    }

    internal static void TextPaint(OfficeDrawing drawing, byte[] bytes, string value, OfficeFontStyle? requiredStyle = null) {
        OfficeDrawingText[] text = DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingText>()
            .Where(item => item.Text == value).ToArray();
        if (text.Length > 0) {
            if (requiredStyle.HasValue) Assert.Contains(text, item => (item.Font.Style & requiredStyle.Value) == requiredStyle.Value);
            return;
        }
        // Embedded PDF faces export exact glyph contours with their accessible source text.
        // A label alone is insufficient: the same group must retain painted geometry.
        XElement group = Assert.Single(XDocument.Parse(System.Text.Encoding.UTF8.GetString(bytes)).Descendants(),
            element => element.Name.LocalName == "g" && element.Attribute("role")?.Value == "img"
                && element.Attribute("aria-label")?.Value.Trim() == value);
        Assert.Contains(group.Descendants(), element => element.Name.LocalName == "path"
            && !string.IsNullOrWhiteSpace(element.Attribute("d")?.Value));
    }

    internal static void WavyStrokes(OfficeDrawing drawing, OfficeColor color, int count) {
        OfficeDrawingShape[] waves = DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingShape>()
            .Where(item => item.Shape.StrokeColor == color && item.Shape.StrokeWidth > 0D
                && item.Shape.FillColor == null && item.Shape.PathCommands.Count >= 4
                && item.Shape.PathCommands.All(command => command.Kind is OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo)
                && item.Shape.Width > item.Shape.Height * 3D
                && item.Shape.Height > item.Shape.StrokeWidth / 2D)
            .ToArray();
        Assert.Equal(count, waves.Length);
        Assert.All(waves, item => {
            OfficePoint[] points = item.Shape.PathCommands.Select(command => command.Point).ToArray();
            Assert.True(points.Zip(points.Skip(1), (first, second) => second.X > first.X).All(increasing => increasing));
            double[] slopes = points.Zip(points.Skip(1), (first, second) => second.Y - first.Y).ToArray();
            Assert.True(slopes.Zip(slopes.Skip(1), (first, second) => first * second < 0D).All(alternating => alternating));
        });
    }

    internal static void DoubleStroke(OfficeDrawing drawing, OfficeColor color) {
        var strokes = DrawingTestTraversal.Elements(drawing).OfType<OfficeDrawingShape>()
            .Where(item => item.Shape.StrokeColor == color && item.Shape.StrokeWidth > 0D && item.Shape.FillColor == null)
            .Select(item => (Item: item, Points: item.Shape.Kind == OfficeShapeKind.Line
                ? item.Shape.Points.ToArray()
                : item.Shape.PathCommands.Where(command => command.Kind is OfficePathCommandKind.MoveTo or OfficePathCommandKind.LineTo)
                    .Select(command => command.Point).ToArray()))
            .Where(item => item.Points.Length == 2 && Math.Abs(item.Points[0].Y - item.Points[1].Y) < 0.001D)
            .ToArray();
        Assert.Contains(strokes, first => strokes.Any(second => !ReferenceEquals(first.Item, second.Item)
            && Math.Abs(first.Item.X - second.Item.X) < 0.001D
            && Math.Abs(first.Item.Shape.Width - second.Item.Shape.Width) < 0.001D
            && Math.Abs(first.Item.Y - second.Item.Y) > 0.1D
            && Math.Abs(first.Item.Y - second.Item.Y) < Math.Max(first.Item.Shape.StrokeWidth, second.Item.Shape.StrokeWidth) * 3D));
    }
}
