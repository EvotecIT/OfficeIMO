using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextGrowthTests {
    [Theory]
    [InlineData(OfficeTextAlignment.Left)]
    [InlineData(OfficeTextAlignment.Center)]
    [InlineData(OfficeTextAlignment.Right)]
    public void NaturalWidthContainsAlignedInkAcrossDrawingCopies(OfficeTextAlignment alignment) {
        var paragraphs = new[] { Paragraph("ABCD", alignment) };
        var metrics = Metrics();
        double width = OfficeDrawingTextLayout.RequiredParagraphFrameWidth(paragraphs, metrics, default);
        Assert.Equal(45, width);
        var original = new OfficeDrawing(width, 24).AddRichTextParagraphs(paragraphs, 0, 0, width, 24);
        Assert.Single(original.Elements.OfType<OfficeDrawingRichText>()).NormalizeHorizontalPaint = true;
        var nested = new OfficeDrawing(width + 20, 44).AddDrawing(original, 10, 10);
        var tinted = original.Clone(); tinted.ApplyColorTint(OfficeColor.FromRgb(20, 60, 80));
        foreach (var drawing in new[] { original, original.Clone(), nested, tinted }) {
            var text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
            var layout = OfficeDrawingTextLayout.CreateWithMetrics(text, width, 24, metrics);
            var paint = OfficeDrawingTextLayout.HorizontalPaintEnvelope(layout, metrics.MeasureHorizontalPaintBounds);
            Assert.Equal(3, Assert.Single(layout.Lines).OffsetX);
            Assert.Equal(0, paint.Left); Assert.Equal(width, paint.Right);
        }
    }

    [Theory]
    [InlineData(OfficeTextAlignment.Left, OfficeTextAreaAlignment.FullWidth, 0, 45)]
    [InlineData(OfficeTextAlignment.Center, OfficeTextAreaAlignment.FullWidth, -7.5, 37.5)]
    [InlineData(OfficeTextAlignment.Right, OfficeTextAreaAlignment.FullWidth, -15, 30)]
    [InlineData(OfficeTextAlignment.Left, OfficeTextAreaAlignment.Center, -7.5, 37.5)]
    [InlineData(OfficeTextAlignment.Left, OfficeTextAreaAlignment.Right, -15, 30)]
    public void CappedUnwrappedInkRetainsParagraphAndAreaAnchor(OfficeTextAlignment alignment,
        OfficeTextAreaAlignment area, double expectedLeft, double expectedRight) {
        var drawing = new OfficeDrawing(30, 24).AddRichTextParagraphs(new[] { Paragraph("ABCD", alignment) },
            0, 0, 30, 24, area, wrapText: false);
        var text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>()); text.NormalizeHorizontalPaint = true;
        var metrics = Metrics();
        foreach (var layout in new[] { OfficeDrawingTextLayout.CreateWithMetrics(text, 30, 24, metrics),
            OfficeDrawingTextLayout.CreateForProjectionWithMetrics(text, 30, 24, metrics) }) {
            var line = Assert.Single(layout.Lines);
            Assert.Equal("ABCD", string.Concat(line.Segments.Select(s => s.Text)));
            Assert.Equal(expectedLeft, line.OffsetX - 3);
            Assert.Equal(expectedRight, line.OffsetX + 42);
        }
    }

    [Fact]
    public void WrappedJustificationReservesInkBeforeDistributingWhitespace() {
        var drawing = new OfficeDrawing(65, 100).AddRichTextParagraphs(new[] { Paragraph("AB CD EF GH", OfficeTextAlignment.Justify) },
            0, 0, 65, 100);
        var text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>()); text.NormalizeHorizontalPaint = true;
        var metrics = Metrics(); var layout = OfficeDrawingTextLayout.CreateWithMetrics(text, 65, 100, metrics);
        Assert.True(layout.Lines.Count > 1);
        var paint = OfficeDrawingTextLayout.HorizontalPaintEnvelope(layout, metrics.MeasureHorizontalPaintBounds);
        Assert.Equal(0, paint.Left); Assert.Equal(65, paint.Right); Assert.False(layout.Clipped);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SoftWrapReservesNewLeadingAndTrailingInkBeforeHeightClipping(bool trailing) {
        var metrics = new OfficeDrawingTextMetrics((body, _, _, _) => (body?.Length ?? 0) * 10D,
            (_, _, _, _) => new OfficeTextPaintBounds(-8, 2),
            (body, _, _, _) => (trailing || body?.StartsWith("f", StringComparison.Ordinal) != true ? 0D : -3D,
                (body?.Length ?? 0) * 10D + (trailing && body?.EndsWith("f", StringComparison.Ordinal) == true ? 3D : 0D)));
        // The interior f has no bearing at either unwrapped outer edge.
        var paragraphs = new[] { Paragraph("A f Z", trailing ? OfficeTextAlignment.Right : OfficeTextAlignment.Left) };
        double natural = OfficeDrawingTextLayout.RequiredParagraphFrameWidth(paragraphs, metrics, default);
        Assert.Equal(50, natural);
        var drawing = new OfficeDrawing(20, 100).AddRichTextParagraphs(paragraphs, 0, 0, 20, 100);
        var text = Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>()); text.NormalizeHorizontalPaint = true;
        var complete = OfficeDrawingTextLayout.CreateWithMetrics(text, 20, 100, metrics);
        Assert.Equal(3, complete.Lines.Count);
        var paint = OfficeDrawingTextLayout.HorizontalPaintEnvelope(complete, metrics.MeasureHorizontalPaintBounds);
        Assert.InRange(paint.Left, 0, 20); Assert.InRange(paint.Right, 0, 20); Assert.False(complete.Clipped);
        double height = OfficeDrawingTextLayout.RequiredFrameHeight(complete, metrics.MeasurePaintBounds);
        Assert.False(OfficeDrawingTextLayout.CreateWithMetrics(text, 20, height, metrics).Clipped);
        var shortFrame = OfficeDrawingTextLayout.CreateWithMetrics(text, 20, 15, metrics);
        Assert.Single(shortFrame.Lines); Assert.True(shortFrame.Clipped);
        // Even a hidden f must reserve its bearing; clipping cannot move kept lines.
        Assert.Equal(complete.Lines[0].OffsetX, shortFrame.Lines[0].OffsetX);
    }

    private static OfficeRichTextParagraph Paragraph(string body, OfficeTextAlignment alignment) =>
        new(new[] { new OfficeRichTextRun(body, 12, OfficeColor.Black) }, alignment);

    private static OfficeDrawingTextMetrics Metrics() => new((text, _, _, _) => (text?.Length ?? 0) * 10D,
        (_, _, _, _) => new OfficeTextPaintBounds(-8, 2),
        (text, _, _, _) => (-3D, (text?.Length ?? 0) * 10D + 2D));
}
