using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("<mi>x</mi>", "x", "\U0001D465")]
    [InlineData("<mi>h</mi>", "h", "\u210E")]
    [InlineData("<mi>π</mi>", "π", "\U0001D70B")]
    [InlineData("<mi>cos</mi>", "cos", "cos")]
    [InlineData("<mi>x\u0301</mi>", "x\u0301", "x\u0301")]
    [InlineData("<mi>∞</mi>", "∞", "∞")]
    [InlineData("<mi mathvariant='normal'>x</mi>", "x", "x")]
    [InlineData("<mi style='text-transform:none'>x</mi>", "x", "x")]
    [InlineData("<mi style='text-transform:initial'>x</mi>", "x", "x")]
    [InlineData("<mi style='text-transform:revert'>x</mi>", "x", "\U0001D465")]
    [InlineData("<mi mathvariant='normal' style='text-transform:math-auto'>x</mi>", "x", "\U0001D465")]
    [InlineData("<mi mathvariant='normal' style='text-transform:revert'>x</mi>", "x", "\U0001D465")]
    [InlineData("<mi>x<!-- split -->h</mi>", "xh", "\U0001D465\u210E")]
    [InlineData("<mi>\U0001D465</mi>", "\U0001D465", "\U0001D465")]
    [InlineData("<mtext style='text-transform:math-auto'>x</mtext>", "x", "\U0001D465")]
    public void HtmlMathMl_MathAutoUsesUnicodeGlyphsAndKeepsLogicalText(string markup, string logical, string painted) {
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false, Margins = HtmlRenderMargins.All(0D) };
        options.Fonts.Add("math", ManagedTextShapingTestAssets.CreateFont('x', 'h', 'c', 'o', 's', 0x03C0,
            0x0301, 0x221E, 0x1D465, 0x210E, 0x1D70B));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<math>" + markup + "</math>", options);
        HtmlRenderDrawing math = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        OfficeDrawingText text = Assert.Single(math.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(logical, text.Text);
        Assert.Equal(painted, text.RasterText);
        Assert.Equal(logical, rendered.Text);
        Assert.Equal(OfficeFontStyle.Regular, text.Font.Style);
    }
    [Fact]
    public void HtmlMathMl_MathAutoHonorsPerTokenOverridesAndIgnoresAnnotations() {
        const string html = "<style>.upright{text-transform:none}.mapped{font-family:math;text-transform:math-auto}</style>"
            + "<math><mrow style='text-transform:none'><mi>x</mi><mi class='upright'>x</mi>"
            + "<mi style='text-transform:inherit'>x</mi><mi mathvariant='normal' class='mapped'>x</mi>"
            + "<semantics><mi>x</mi><annotation encoding='text/plain'>description</annotation></semantics></mrow></math>";
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("math", ManagedTextShapingTestAssets.CreateFont('x', 0x1D465));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        var tokens = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>()).Drawing.Elements.OfType<OfficeDrawingText>().ToArray();
        Assert.Equal(new[] { "x", "x", "x", "x", "x" }, tokens.Select(t => t.Text));
        Assert.Equal(new[] { "\U0001D465", "x", "x", "\U0001D465", "\U0001D465" }, tokens.Select(t => t.RasterText));
        Assert.Equal("xxxxx", rendered.Text);
    }

    [Theory]
    [InlineData("x", "\U0001D465")]
    [InlineData("xy", "xy")]
    [InlineData("x<span>h</span>", "\U0001D465\u210E")]
    public void HtmlText_MathAutoUsesTextNodeBoundaries(string markup, string expected) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<p style='text-transform:math-auto'>" + markup + "</p>");
        Assert.Equal(expected, string.Concat(rendered.Pages[0].Visuals.OfType<HtmlRenderText>().Select(x => x.Text)));
    }

    [Fact]
    public void HtmlMathMl_MathAutoPdfRetainsSearchableLogicalTextAndOriginalSupplement() {
        const string html = "<math xmlns='http://www.w3.org/1998/Math/MathML' aria-label='x over two'>"
            + "<mfrac><mi>x</mi><mn>2</mn></mfrac></math>";
        var options = new HtmlToPdfOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("math", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('x', '2', 0x1D465));
        var document = HtmlConversionDocument.Parse(html);
        byte[] pdf = document.ToPdfBytes(options);
        Assert.Contains("(x)/(2)", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
        var rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderSemanticGroup formula = Assert.Single(rendered.Pages.SelectMany(page => EnumerateMathMlScene(page.Scene)).OfType<HtmlRenderSemanticGroup>(),
            x => x.Role == HtmlRenderSemanticGroupRole.Formula);
        Assert.Equal(html, formula.MathMlSource!.MathMl);
        Assert.True(formula.MathMlSource.IsOriginalMarkup);
    }

    [Fact]
    public void HtmlMathMl_MathAutoRetainsIdentifierGlyphsInFunctionsAndAccents() {
        const string html = "<math><mrow><mi>f</mi><mo>⁡</mo><mfenced><mi>x</mi></mfenced></mrow></math>"
            + "<math><mover accent='true'><mi>x</mi><mi>h</mi></mover></math>";
        var options = new HtmlRenderOptions { AllowSystemFontFallback = false };
        options.Fonts.Add("math", ManagedTextShapingTestAssets.CreateFont('f', 'x', 'h', '(', ')', 0x1D453, 0x1D465, 0x210E));
        HtmlRenderDrawing[] drawings = HtmlRenderTestDriver.Render(html, options).Pages[0].Visuals.OfType<HtmlRenderDrawing>().ToArray();
        Assert.Equal(new[] { "\U0001D453", "(", "\U0001D465", ")" }, drawings[0].Drawing.Elements.OfType<OfficeDrawingText>().Select(t => t.RasterText));
        Assert.Equal("\U0001D465", Assert.Single(drawings[1].Drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "x").RasterText);
        Assert.Equal("\u210E", Assert.Single(drawings[1].Drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "h").RasterText);
    }

}
