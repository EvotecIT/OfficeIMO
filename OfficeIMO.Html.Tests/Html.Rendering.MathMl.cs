using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMathMl_RendersStructuredInlineMathWithMeasuredBaselineAndLogicalText() {
        const string html = "<body style='margin:0'><p style='margin:0;font-size:20px;line-height:24px'>Before "
            + "<math id='fraction' aria-label='x divided by two'><mfrac><mi>x</mi><mn>2</mn></mfrac></math> after</p></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 260D,
            ViewportHeight = 100D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderLogicalTextGroup logical = Assert.Single(
            EnumerateMathMlScene(rendered.Pages[0].Scene).OfType<HtmlRenderLogicalTextGroup>(),
            item => item.Source == "math#fraction");
        HtmlRenderDrawing drawing = Assert.Single(logical.Visuals.OfType<HtmlRenderDrawing>());
        HtmlRenderText before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text.Contains("Before", StringComparison.Ordinal));

        Assert.Equal("(x)/(2)", logical.Text);
        Assert.Equal("x divided by two", drawing.AlternativeText);
        Assert.Equal("Before \n(x)/(2)\n after", rendered.Text);
        Assert.Contains(drawing.Drawing.Elements, item => item is OfficeDrawingShape);
        Assert.Contains(drawing.Drawing.Elements, item => item is OfficeDrawingText text && text.Text == "x");
        Assert.Contains(drawing.Drawing.Elements, item => item is OfficeDrawingText text && text.Text == "2");
        Assert.True(drawing.Y < before.Y, $"Expected the fraction to rise above the neighboring text, but math Y={drawing.Y} and text Y={before.Y}.");
        Assert.True(drawing.Y + drawing.Height > before.Y + before.Font.Size, "Expected the fraction denominator to extend below the neighboring text baseline.");
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
    }

    [Fact]
    public void HtmlMathMl_UsesSharedAccessibleNameResolutionBeforeAltText() {
        const string html = "<span id='math-name'>Quadratic expression</span>"
            + "<math id='formula' aria-labelledby='math-name' alttext='Lower-priority alternative'><mi>x</mi></math>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderDrawing drawing = Assert.Single(
            EnumerateMathMlScene(rendered.Pages[0].Scene).OfType<HtmlRenderDrawing>(),
            item => item.Source == "math#formula");

        Assert.Equal("Quadratic expression", drawing.AlternativeText);
    }

    [Fact]
    public void HtmlMathMl_HonorsBlockDisplayAndAuthoredSizeAcrossSvgAndSearchablePdf() {
        const string html = "<body style='margin:0'><div style='font-size:12px;line-height:14px'>Above</div>"
            + "<math id='root' display='block' aria-label='square root of x' style='width:120px;height:48px'>"
            + "<msqrt><mi>x</mi></msqrt></math><div style='font-size:12px;line-height:14px'>Below</div></body>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(180D / HtmlRenderOptions.CssPixelsPerInch, 120D / HtmlRenderOptions.CssPixelsPerInch),
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source, options);
        HtmlRenderDrawing drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>(), item => item.Source == "math#root");
        HtmlRenderText above = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Above");
        HtmlRenderText below = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Below");
        string svg = source.ToSvg(options);
        byte[] pdf = source.ToPdf(new HtmlPdfSaveOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = new OfficePageSize(180D / HtmlRenderOptions.CssPixelsPerInch, 120D / HtmlRenderOptions.CssPixelsPerInch),
            HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(120D, drawing.Width, 3);
        Assert.Equal(48D, drawing.Height, 3);
        Assert.True(drawing.Y >= above.Y + above.Height - 0.01D);
        Assert.True(below.Y >= drawing.Y + drawing.Height - 0.01D);
        Assert.Contains("sqrt(x)", rendered.Text, StringComparison.Ordinal);
        Assert.Contains(">x</text>", svg, StringComparison.Ordinal);
        Assert.Contains("sqrt(x)", pdfText, StringComparison.Ordinal);
        Assert.Equal(pdfText.IndexOf("sqrt(x)", StringComparison.Ordinal), pdfText.LastIndexOf("sqrt(x)", StringComparison.Ordinal));
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
    }

    [Fact]
    public void HtmlMathMl_UsesDiagnosedTextFallbackForMalformedPresentationStructure() {
        const string html = "<body style='margin:0'><math id='broken'><mfrac><mi>x</mi></mfrac></math></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Contains("x", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Pages[0].Visuals, item => item is HtmlRenderDrawing drawing && drawing.Source == "math#broken");
    }

    [Fact]
    public void HtmlMathMl_DiagnosesUnsupportedElementsWhileRetainingChildContent() {
        const string html = "<body style='margin:0'><math id='action'><maction><mi>x</mi></maction></math></body>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);

        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
        Assert.Equal("maction", diagnostic.Detail);
        Assert.Equal("x", rendered.Text);
        Assert.Contains(rendered.Pages[0].Visuals, item => item is HtmlRenderDrawing drawing && drawing.Source == "math#action");
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateMathMlScene(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderFormField form => form.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateMathMlScene(children)) yield return child;
        }
    }
}
