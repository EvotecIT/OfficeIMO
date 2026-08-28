using System;
using System.Linq;
using System.Reflection;
using System.Text;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeRichTextFormattingTests {
    [Fact]
    public void PreTypographyConstructorsRemainBinaryDiscoverable() {
        ConstructorInfo? run = typeof(OfficeRichTextRun).GetConstructor(new[] {
            typeof(string), typeof(double), typeof(OfficeColor), typeof(bool), typeof(bool),
            typeof(bool), typeof(string), typeof(bool), typeof(OfficeColor?)
        });
        ConstructorInfo? segment = typeof(OfficeRichTextSegment).GetConstructor(new[] {
            typeof(string), typeof(double), typeof(double), typeof(OfficeColor), typeof(bool),
            typeof(bool), typeof(bool), typeof(string), typeof(bool), typeof(OfficeColor?)
        });

        Assert.NotNull(run);
        Assert.NotNull(segment);
    }

    [Fact]
    public void PreTypographyRendererMethodsRemainBinaryDiscoverable() {
        Assert.NotNull(typeof(OfficeRasterCanvas).GetMethod(nameof(OfficeRasterCanvas.DrawTextLine), new[] {
            typeof(string), typeof(double), typeof(double), typeof(double), typeof(OfficeColor),
            typeof(bool), typeof(bool), typeof(OfficeTextAlignment), typeof(double), typeof(double), typeof(double),
            typeof(bool), typeof(bool), typeof(string), typeof(bool), typeof(bool)
        }));
        AssertMethod(nameof(OfficeTextBlockRenderer.DrawRasterTextBlock),
            typeof(OfficeRasterCanvas), typeof(OfficeTextBlockLayout),
            typeof(double), typeof(double), typeof(double), typeof(double), typeof(OfficeColor),
            typeof(OfficeTextAlignment), typeof(OfficeTextVerticalAlignment),
            typeof(bool), typeof(bool), typeof(bool),
            typeof(double), typeof(double), typeof(double), typeof(bool), typeof(double), typeof(bool), typeof(string),
            typeof(bool), typeof(bool));
        AssertMethod(nameof(OfficeTextBlockRenderer.DrawRasterTextBox),
            typeof(OfficeRasterCanvas), typeof(OfficeTextBlockRenderPlan), typeof(OfficeColor),
            typeof(bool), typeof(bool), typeof(bool), typeof(OfficeTextAlignment?), typeof(OfficeTextVerticalAlignment?),
            typeof(double), typeof(double), typeof(double), typeof(OfficeColor?), typeof(double), typeof(double),
            typeof(bool), typeof(double), typeof(bool), typeof(string));
        AssertMethod(nameof(OfficeTextBlockRenderer.AppendSvgTextElement),
            typeof(StringBuilder), typeof(string), typeof(double), typeof(double), typeof(double), typeof(OfficeColor),
            typeof(string), typeof(double), typeof(OfficeTextAlignment), typeof(bool), typeof(bool), typeof(bool),
            typeof(double), typeof(double), typeof(double), typeof(bool));
        AssertMethod(nameof(OfficeTextBlockRenderer.WriteSvgTextBlock),
            typeof(System.Xml.XmlWriter), typeof(OfficeTextBlockLayout),
            typeof(double), typeof(double), typeof(double), typeof(double), typeof(OfficeColor), typeof(string),
            typeof(OfficeTextAlignment), typeof(OfficeTextVerticalAlignment), typeof(bool), typeof(bool), typeof(bool),
            typeof(double), typeof(double), typeof(double), typeof(string), typeof(Action<System.Xml.XmlWriter>), typeof(bool));
        AssertMethod(nameof(OfficeTextBlockRenderer.WriteSvgTextBox),
            typeof(System.Xml.XmlWriter), typeof(OfficeTextBlockRenderPlan), typeof(OfficeColor), typeof(string),
            typeof(bool), typeof(bool), typeof(bool), typeof(double), typeof(double), typeof(double), typeof(string),
            typeof(OfficeColor?), typeof(double), typeof(double),
            typeof(Action<System.Xml.XmlWriter>), typeof(Action<System.Xml.XmlWriter>), typeof(bool));
    }

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
    public void ScriptEllipsisUsesTheRenderedFontSizeWhenNoSourceTextFits() {
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutRichTextBlock(
            new[] { new OfficeRichTextRun("WWWW", 10D, OfficeColor.Black, baseline: OfficeTextBaseline.Superscript) },
            maxWidth: 19.5D,
            maxHeight: 20D,
            lineHeightFactor: 1D,
            measure: static (text, size, _) => (text?.Length ?? 0) * size,
            wrap: false);

        OfficeRichTextSegment segment = Assert.Single(Assert.Single(layout.Lines).Segments);
        Assert.Equal("...", segment.Text);
        Assert.Equal(19.5D, segment.Width);
        Assert.Equal(OfficeTextBaseline.Superscript, segment.Baseline);
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
    public void SvgKeepsIndependentUnderlineAndStrikethroughPatterns() {
        var layout = new OfficeTextBlockLayout(
            new[] { new OfficeTextLine("Styled", 60D) },
            fontSize: 16D,
            lineHeight: 20D,
            width: 60D,
            height: 20D);
        var builder = new StringBuilder();

        builder.AppendSvgStyledTextBlock(
            layout, 0D, 0D, 100D, 30D, OfficeColor.Black,
            "Aptos", OfficeTextAlignment.Left, OfficeTextVerticalAlignment.Top,
            bold: false, italic: false, underline: true,
            rotationDegrees: 0D,
            rotationCenterX: 0D,
            rotationCenterY: 0D,
            centerLineInLineHeight: true,
            underlineStyle: OfficeTextDecorationStyle.Double,
            strikethrough: true,
            strikethroughStyle: OfficeTextDecorationStyle.Dotted,
            baseline: OfficeTextBaseline.Normal);

        string svg = builder.ToString();
        Assert.Contains("text-decoration=\"line-through\"", svg, StringComparison.Ordinal);
        Assert.Contains("text-decoration-style=\"dotted\"", svg, StringComparison.Ordinal);
        Assert.Contains("<tspan text-decoration=\"underline\" text-decoration-style=\"double\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void BaselineTextBlockLayoutsUseTheRenderedFontSize() {
        var drawing = new OfficeDrawing(80D, 30D)
            .AddStyledText(
                "123456", 0D, 0D, 40D, 24D,
                new OfficeFontInfo("Aptos", 20D, OfficeFontStyle.Regular), OfficeColor.Black,
                OfficeTextAlignment.Left, null, OfficeTextVerticalAlignment.Top,
                0D, null, null,
                true, false, false, false, false,
                null, null,
                OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None,
                OfficeTextBaseline.Superscript);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Contains("font-size=\"13\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("font-size=\"8.45\"", svg, StringComparison.Ordinal);
        Assert.True(CountPaintedPixels(raster) > 0);
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
    public void RasterWavyDecorationDoesNotCollapseToASingleLine() {
        byte[] single = RenderDecoration(OfficeTextDecorationStyle.Single).GetPixels();
        byte[] wavy = RenderDecoration(OfficeTextDecorationStyle.Wavy).GetPixels();

        Assert.False(single.SequenceEqual(wavy));
        Assert.True(CountPaintedRows(wavy, width: 180, minimumY: 27) >= 3);
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
            8D,
            8D,
            24D,
            OfficeColor.Black,
            false,
            false,
            OfficeTextAlignment.Left,
            0D,
            0D,
            0D,
            false,
            false,
            null,
            false,
            false,
            style,
            OfficeTextDecorationStyle.None);
        return image;
    }

    private static void AssertMethod(string name, params Type[] parameters) =>
        Assert.NotNull(typeof(OfficeTextBlockRenderer).GetMethod(name, parameters));

    private static int CountPaintedPixels(OfficeRasterImage image) =>
        image.GetPixels().Where((_, index) => index % 4 == 3).Count(alpha => alpha != 0);

    private static int CountPaintedRows(byte[] pixels, int width, int minimumY) =>
        Enumerable.Range(minimumY, pixels.Length / (width * 4) - minimumY)
            .Count(y => Enumerable.Range(0, width).Any(x => pixels[((y * width) + x) * 4 + 3] != 0));
}
