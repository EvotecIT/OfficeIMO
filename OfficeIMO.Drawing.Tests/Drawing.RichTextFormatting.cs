using System;
using System.Linq;
using System.Text;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeRichTextFormattingTests {
    [Fact]
    public void RichTextRunKeepsCompatibilityBooleansAndCanonicalStyles() {
        var run = new OfficeRichTextRun(
            "Styled",
            20D,
            OfficeColor.Blue,
            underlineStyle: OfficeTextDecorationStyle.Dotted,
            strikethroughStyle: OfficeTextDecorationStyle.Double,
            baseline: OfficeTextBaseline.Superscript);

        Assert.True(run.Underline);
        Assert.True(run.Strikethrough);
        Assert.Equal(OfficeTextDecorationStyle.Dotted, run.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, run.StrikethroughStyle);
        Assert.Equal(OfficeTextBaseline.Superscript, run.Baseline);
        Assert.Equal(13D, run.EffectiveFontSize);

        OfficeRichTextRun copy = run.WithTextCase(OfficeTextCase.ToggleCase);
        Assert.Equal("sTYLED", copy.Text);
        Assert.Equal(run.UnderlineStyle, copy.UnderlineStyle);
        Assert.Equal(run.StrikethroughStyle, copy.StrikethroughStyle);
        Assert.Equal(run.Baseline, copy.Baseline);
    }

    [Fact]
    public void LayoutMeasuresScriptRunsAtTheirRenderedSize() {
        var runs = new[] {
            new OfficeRichTextRun("A", 20D, OfficeColor.Black),
            new OfficeRichTextRun("2", 20D, OfficeColor.Black, baseline: OfficeTextBaseline.Subscript)
        };

        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            runs,
            maxWidth: 100D,
            maxHeight: 40D,
            lineHeightFactor: 1.2D,
            measure: static (text, size, _) => (text?.Length ?? 0) * size,
            wrap: false);

        OfficeRichTextLine line = Assert.Single(layout.Lines);
        Assert.Equal(33D, line.Width);
        Assert.Equal(13D, line.Segments[1].Width);
        Assert.Equal(OfficeTextBaseline.Subscript, line.Segments[1].Baseline);
    }

    [Fact]
    public void SvgPreservesDecorationPatternAndScriptPlacement() {
        var segment = new OfficeRichTextSegment(
            "H2O",
            width: 20D,
            fontSize: 20D,
            color: OfficeColor.FromRgb(12, 34, 56),
            bold: true,
            italic: true,
            underline: false,
            fontFamily: "Aptos",
            underlineStyle: OfficeTextDecorationStyle.Wavy,
            baseline: OfficeTextBaseline.Subscript);
        var builder = new StringBuilder();

        builder.AppendSvgRichTextSegment(segment, 5D, 30D);

        string svg = builder.ToString();
        Assert.Contains("y=\"33\"", svg, StringComparison.Ordinal);
        Assert.Contains("font-size=\"13\"", svg, StringComparison.Ordinal);
        Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
        Assert.Contains("text-decoration-style=\"wavy\"", svg, StringComparison.Ordinal);
        Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
        Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterDistinguishesSingleAndDoubleDecorations() {
        OfficeRasterImage single = RenderDecoration(OfficeTextDecorationStyle.Single);
        OfficeRasterImage doubled = RenderDecoration(OfficeTextDecorationStyle.Double);

        int singlePixels = CountPaintedPixels(single);
        int doublePixels = CountPaintedPixels(doubled);
        Assert.True(singlePixels > 0);
        Assert.True(doublePixels > singlePixels, $"Expected a double underline to paint more pixels than a single underline ({doublePixels} <= {singlePixels}).");
    }

    [Fact]
    public void RichTextRejectsUndefinedFormattingValues() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeRichTextRun(
            "Invalid",
            12D,
            OfficeColor.Black,
            underlineStyle: (OfficeTextDecorationStyle)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeRichTextRun(
            "Invalid",
            12D,
            OfficeColor.Black,
            baseline: (OfficeTextBaseline)99));
    }

    private static OfficeRasterImage RenderDecoration(OfficeTextDecorationStyle style) {
        var image = new OfficeRasterImage(180, 50, OfficeColor.Transparent);
        var canvas = new OfficeRasterCanvas(image);
        canvas.DrawTextLine(
            "Decoration",
            anchorX: 8D,
            top: 8D,
            height: 24D,
            color: OfficeColor.Black,
            alignment: OfficeTextAlignment.Left,
            underlineStyle: style);
        return image;
    }

    private static int CountPaintedPixels(OfficeRasterImage image) =>
        image.GetPixels().Where((_, index) => index % 4 == 3).Count(alpha => alpha != 0);
}
