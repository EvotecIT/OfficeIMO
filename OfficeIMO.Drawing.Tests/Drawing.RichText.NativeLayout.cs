using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingNativeTextLayoutTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ColorFontFitIncludesThePaintedMonochromeOutline(bool rich) {
        var drawing = new OfficeDrawing(100, 50);
        drawing.Fonts.Add("Color Fit", OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateColorFont('A', baseGlyphHeight: 1000));
        if (rich) drawing.AddRichText(new[] { new OfficeRichTextRun("A", 20, OfficeColor.Black, fontFamily: "Color Fit") },
            10, 10, 80, 15, lineHeight: 10, shrinkToFit: true);
        else drawing.AddText("A", 10, 10, 80, 15, new OfficeFontInfo("Color Fit", 20),
            lineHeight: 10, wrapText: true, shrinkToFit: true);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 3, OfficeColor.White);
        int ink = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
                if (image.GetPixel(x, y).R < 160) {
                    ink++;
                    Assert.InRange(y, 30, 74);
                }
        Assert.True(ink > 20);
    }

    [Theory]
    [InlineData(false, "gypsy", OfficeTextVerticalAlignment.Top)]
    [InlineData(true, "gypsy", OfficeTextVerticalAlignment.Bottom)]
    [InlineData(false, "\u00C1gj", OfficeTextVerticalAlignment.Center)]
    [InlineData(true, "\u00C1gj", OfficeTextVerticalAlignment.Bottom)]
    public void FittedGlyphPixelsStayInsideTheActualFrame(bool rich, string value, OfficeTextVerticalAlignment alignment) {
        var drawing = new OfficeDrawing(100, 50);
        drawing.Fonts.Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        if (rich) drawing.AddRichText(new[] { new OfficeRichTextRun(value, 20, OfficeColor.Black, fontFamily: "Proof Sans") },
            10, 10, 80, 15, lineHeight: 10, verticalAlignment: alignment, shrinkToFit: true);
        else drawing.AddText(value, 10, 10, 80, 15, new OfficeFontInfo("Proof Sans", 20),
            lineHeight: 10, verticalAlignment: alignment, wrapText: true, shrinkToFit: true);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 3, OfficeColor.White);
        int ink = 0;
        for (int y = 0; y < image.Height; y++)
            for (int x = 0; x < image.Width; x++)
                if (image.GetPixel(x, y).R < 160) {
                    ink++;
                    Assert.InRange(y, 30, 74);
                }
        Assert.True(ink > 20);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterExportScalePreservesMinimumFontSize(bool stacked) {
        var drawing = new OfficeDrawing(100, 30).AddText(stacked ? "g" : "gypsy", 10, 10, 80, 4,
            new OfficeFontInfo("Proof Sans", 20), lineHeight: 10, wrapText: true, stackedText: stacked, shrinkToFit: true);
        drawing.Fonts.Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        int InkWidth(double scale) {
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, scale, OfficeColor.White);
            int first = image.Width, last = -1;
            for (int y = 0; y < image.Height; y++)
                for (int x = 0; x < image.Width; x++)
                    if (image.GetPixel(x, y).R < 160) { first = Math.Min(first, x); last = Math.Max(last, x); }
            Assert.True(last >= first);
            return last - first + 1;
        }
        Assert.InRange(InkWidth(2), InkWidth(1) * 2 - 1, InkWidth(1) * 2 + 1);
    }

    [Theory]
    [InlineData("rich")]
    [InlineData("wrapped")]
    [InlineData("stacked")]
    [InlineData("stacked-rich")]
    public void MinimumFontStillReportsPaintThatCannotFit(string mode) {
        double Measure(string? value, double size) => (value?.Length ?? 0) * size / 2;
        var runs = new[] { new OfficeRichTextRun("g", 20, OfficeColor.Red) };
        if (mode == "rich" || mode == "stacked-rich") {
            var layout = mode == "rich"
                ? OfficeDrawingTextLayout.Create(new OfficeDrawingRichText(runs, 0, 0, 100, 4, lineHeight: 10, shrinkToFit: true),
                    100, 4, (value, size, family, style) => Measure(value, size))
                : OfficeTextLayoutEngine.LayoutStackedRichTextBlock(runs, 100, 4, .5D, Measure, minimumFontSize: 6);
            Assert.True(layout.Clipped);
            Assert.Equal(6D, Assert.Single(Assert.Single(layout.Lines).Segments).FontSize);
        } else {
            var layout = mode == "wrapped"
                ? OfficeTextLayoutEngine.FitWrappedText("g", 20, 100, 4, .5D, 6, Measure)
                : OfficeTextLayoutEngine.LayoutStackedTextBlock("g", 20, 100, 4, .5D, 6, Measure);
            Assert.True(layout.Clipped);
            Assert.Equal(6D, layout.FontSize);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CondensedPlainFrameFitIncludesGlyphHeight(bool stacked) {
        double Measure(string? value, double size) => (value?.Length ?? 0) * size / 2;
        var layout = stacked
            ? OfficeTextLayoutEngine.LayoutStackedTextBlock("g", 20, 100, 15, .5D, 6, Measure)
            : OfficeTextLayoutEngine.FitWrappedText("gypsy", 20, 100, 15, .5D, 6, Measure);
        double top = OfficeTextPlacement.ResolveTop(0, 15, layout.Height, OfficeTextVerticalAlignment.Bottom);
        Assert.InRange(top + OfficeDrawingTextLayout.PaintedHeight(layout), 0, 15.01D);
        Assert.False(layout.Clipped);
    }

    [Fact]
    public void CondensedStackedRichFrameFitIncludesGlyphHeight() {
        var layout = OfficeTextLayoutEngine.LayoutStackedRichTextBlock(
            new[] { new OfficeRichTextRun("g", 20, OfficeColor.Red) }, 100, 15, .5D,
            (value, size) => (value?.Length ?? 0) * size / 2, minimumFontSize: 6);
        Assert.InRange(OfficeDrawingTextLayout.PaintedHeight(layout), 0, 15.01D);
        Assert.False(layout.Clipped);
    }

    [Theory]
    [InlineData(OfficeTextVerticalAlignment.Top, 1D)]
    [InlineData(OfficeTextVerticalAlignment.Center, 1D)]
    [InlineData(OfficeTextVerticalAlignment.Bottom, 1D)]
    [InlineData(OfficeTextVerticalAlignment.Top, 2D)]
    public void CondensedRichTextFitsItsPaintedHeight(OfficeTextVerticalAlignment alignment, double scale) {
        var text = new OfficeDrawingRichText(new[] { new OfficeRichTextRun("gypsy", 20, OfficeColor.Red) },
            0, 0, 100, 15, lineHeight: 10, verticalAlignment: alignment, shrinkToFit: true);
        var layout = OfficeDrawingTextLayout.Create(text, 100 * scale, 15 * scale,
            (value, size, family, style) => (value?.Length ?? 0) * size / 2, scale);
        double top = OfficeTextPlacement.ResolveTop(0, 15 * scale, layout.Height, alignment);
        Assert.InRange(top + OfficeDrawingTextLayout.PaintedHeight(layout), 0, 15 * scale + .01D);
        Assert.False(layout.Clipped);
        Assert.Equal("gypsy", string.Concat(layout.Lines.SelectMany(line => line.Segments).Select(segment => segment.Text)));
    }

    [Theory]
    [InlineData(1D, 9D)]
    [InlineData(2D, 18D)]
    public void DrawingRichTextRetainsCondensedLeading(double scale, double expectedLineHeight) {
        var drawing = new OfficeDrawing(100, 40).AddRichText(new[] {
            new OfficeRichTextRun("first\nsecond\nthird", 12D, OfficeColor.Black)
        }, 0, 0, 100, 40, lineHeight: 9D);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.Create(text, 100 * scale, 40 * scale,
            (value, size, family, style) => (value?.Length ?? 0) * size / 2, scale);
        Assert.Equal(3, layout.Lines.Count);
        Assert.Equal(expectedLineHeight, layout.LineHeight);
        Assert.All(layout.Lines, line => Assert.Equal(expectedLineHeight, line.LineHeight));
        Assert.Equal(3 * expectedLineHeight, layout.Height);
        Assert.False(layout.Clipped);
    }

    [Theory]
    [InlineData("alpha beta gamma delta", true, 60D, 25D)]
    [InlineData("first\nsecond\nthird", false, 100D, 27D)]
    [InlineData("x\na considerably longer line", false, 80D, 25D)]
    public void DrawingRichTextShrinksAgainstWidthAndHeight(string value, bool wrap, double width, double height) {
        var drawing = new OfficeDrawing(width, height).AddRichText(new[] {
            new OfficeRichTextRun(value, 18D, OfficeColor.Red, bold: true) { LinkUri = "https://officeimo.net/fit" }
        }, 0, 0, width, height, wrapText: wrap, shrinkToFit: true);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.Create(text, width, height,
            (content, size, family, style) => (content?.Length ?? 0) * size / 2);
        Assert.False(layout.Clipped);
        Assert.InRange(layout.Width, 0D, width + 0.01D);
        Assert.InRange(layout.Height, 0D, height + 0.01D);
        var segments = layout.Lines.SelectMany(line => line.Segments).ToArray();
        Assert.Equal(value.Replace(" ", "").Replace("\n", ""), string.Concat(segments.Select(segment => segment.Text)).Replace(" ", ""));
        Assert.All(segments, segment => {
            Assert.InRange(segment.FontSize, 6D, 17.999D);
            Assert.True(segment.Bold);
            Assert.Equal(OfficeColor.Red, segment.Color);
            Assert.Equal("https://officeimo.net/fit", segment.LinkUri);
        });
    }

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
