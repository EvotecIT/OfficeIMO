using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNativeTextLayoutTests {
    [Theory]
    [InlineData(1D, 10D)]
    [InlineData(2D, 20D)]
    public void SmallRichTextPreservesAuthoredLineSpacing(double scale, double expectedLineHeight) {
        var drawing = new OfficeDrawing(100, 50).AddRichText(new[] {
            new OfficeRichTextRun("first\nsecond", 8D, OfficeColor.Black)
        }, 0, 0, 100, 50, lineHeight: 9.6D);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.Create(text, 100 * scale, 50 * scale,
            (value, size, family, style) => (value?.Length ?? 0) * size / 2, scale);
        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal(expectedLineHeight, layout.LineHeight);
        Assert.Equal(2 * expectedLineHeight, layout.Height);
    }

    [Fact]
    public void UnboundedHeightMeasurementPreservesAllWrappedLines() {
        var layout = OfficeTextLayoutEngine.LayoutTextBlock("first second third", 12, 40,
            double.MaxValue, 1.2D, 6, (text, size) => (text?.Length ?? 0) * 6D, wrap: true);
        Assert.Equal(new[] { "first", "second", "third" }, layout.Lines.Select(line => line.Text));
        Assert.Equal(45D, layout.Height);
        Assert.False(layout.Clipped);
    }

    [Theory]
    [InlineData("a b", 25D)]
    [InlineData("a b c", 45D)]
    public void WrappingUsesTheMeasuredJoinedTextAdvance(string text, double availableWidth) {
        // A font-measurement boundary with a pair adjustment across a token boundary.
        double Measure(string? value, double size) => (value?.Length ?? 0) * 10D - (value?.Contains("a b") == true ? 5D : 0D);
        var layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            new[] { new OfficeRichTextRun(text, 12D, OfficeColor.Black) }, availableWidth, 50D, 1.2D, Measure, true);
        OfficeRichTextLine line = Assert.Single(layout.Lines);
        Assert.Equal(text, string.Concat(line.Segments.Select(segment => segment.Text)));
        Assert.Equal(availableWidth, line.Width);
    }

    [Fact]
    public void StyledWrappingPreservesDistinctLinkTargetsAcrossCloningAndLineBreaks() {
        var drawing = new OfficeDrawing(150, 70).AddRichText(new[] {
            new OfficeRichTextRun("first link ", 12, OfficeColor.Black, bold: true) { LinkUri = "https://officeimo.net/first" },
            new OfficeRichTextRun("second link", 12, OfficeColor.Black, bold: true) { LinkUri = "https://officeimo.net/second" }
        }, 0, 0, 150, 70, wrapText: true);
        OfficeDrawingRichText text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Clone().Elements));
        var layout = OfficeDrawingTextLayout.Create(text, 150, 70,
            (value, size, family, style) => (value?.Length ?? 0) * (style.HasFlag(OfficeFontStyle.Bold) ? 10D : 5D));
        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal("https://officeimo.net/first", Assert.Single(layout.Lines[0].Segments).LinkUri);
        Assert.Equal("https://officeimo.net/second", Assert.Single(layout.Lines[1].Segments).LinkUri);
        Assert.Equal("second link", Assert.Single(layout.Lines[1].Segments).Text);
    }
}
