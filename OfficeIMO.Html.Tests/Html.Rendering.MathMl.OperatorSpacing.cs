using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMathMl_AuthoredCssPixelsKeepFormulaSemanticsAndOriginalSupplement() {
        const string markup = "<math xmlns='http://www.w3.org/1998/Math/MathML' aria-label='spaced x equals two' "
            + "style='font:20px SpacingMath'><mrow><mtext>x</mtext><mo form='infix' lspace='4px' rspace='0'>=</mo><mn>2</mn></mrow></math>";
        var options = new HtmlToPdfOptions { PageSize = new OfficePageSize(3D, 2D), HonorCssPageRules = false,
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false };
        options.Fonts.Add("SpacingMath", ManagedTextShapingTestAssets.CreateFont('x', '=', '2'));
        var result = HtmlConversionDocument.Parse(markup).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        var drawing = Assert.Single(result.RenderResult.Document.Pages.SelectMany(p => p.Visuals).OfType<HtmlRenderDrawing>());
        Assert.Equal(34D, drawing.Drawing.Width, 6);
        Assert.Equal(14D, Assert.Single(drawing.Drawing.Elements.OfType<OfficeDrawingText>(), t => t.Text == "=").X, 6);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        Assert.Equal(1, text.Split(new[] { "x=2" }, StringSplitOptions.None).Length - 1);
        var tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        Assert.Equal("spaced x equals two", Assert.Single(tagged.StructureElements, e => e.StructureType == "Formula").AlternateText);
        Assert.Equal(markup, System.Text.Encoding.UTF8.GetString(Assert.Single(PdfCore.PdfAttachmentExtractor.ExtractAttachments(pdf)).Bytes));
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(pdf));
    }

    [Theory]
    [InlineData("<mtext>x</mtext><mo lspace='2ex'>=</mo><mn>2</mn>", "mo[lspace]")]
    [InlineData("<mtext>x</mtext><mo rspace='calc(1em + 1px)'>=</mo><mn>2</mn>", "mo[rspace]")]
    [InlineData("<msup><mrow><mo lspace='.5em'>+</mo></mrow><mn>2</mn></msup>", "mo[operator-spacing-in-msup]")]
    public void HtmlMathMl_DiagnosesSpacingOutsideTheQualifiedOrdinaryRowSubset(string content, string detail) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render("<math><mrow>" + content + "</mrow></math>");
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.MathMlContentUnsupported);
        Assert.Equal(detail, diagnostic.Detail);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.NotEmpty(rendered.Text);
    }
}
