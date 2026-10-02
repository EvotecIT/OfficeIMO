using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlMathMlPdf_BalancedPolicyUsesMathWithoutSelectingNamedHostFonts() {
        if (!HasInstalledMathematicalFace()) return;
        string? installedFamily = new[] { "Helvetica Neue", "Arial", "DejaVu Sans" }
            .FirstOrDefault(name => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
                "<p style='font-family:" + name + ",system-ui'>x</p>")).Fonts.Faces
                .Any(face => face.FamilyName == name));
        if (installedFamily == null) return;
        string child = "<p style=\"font-family:'" + installedFamily + "',math\">CHILD</p>";
        var source = HtmlConversionDocument.Parse("<math><mtext>x</mtext></math>"
            + "<p style=\"font-family:'" + installedFamily + "',math\">PARENT</p>"
            + "<iframe style='width:200px;height:80px' srcdoc=\""
            + System.Net.WebUtility.HtmlEncode(child) + "\"></iframe>");
        var options = new HtmlToPdfOptions();
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(source, options);
        HtmlRenderDocument rendered = result.RenderResult!.Document;

        Assert.True(rendered.Fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out _));
        Assert.DoesNotContain(rendered.Fonts.Faces, face => face.FamilyName == installedFamily);
        Assert.DoesNotContain(result.RenderResult.Diagnostics, diagnostic =>
            diagnostic.Code == "InstalledFontResolved" && diagnostic.Source == installedFamily);
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText();
        Assert.Contains("PARENT", text, StringComparison.Ordinal);
        Assert.Contains("CHILD", text, StringComparison.Ordinal);
        Assert.Empty(options.Fonts.Faces);
        Assert.False(options.ResourcePolicy.AllowDocumentFontEmbedding);
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void HtmlMathMlPdf_DisabledSystemEmbeddingOrFallbackDoesNotLoadMath(bool systemFonts, bool fallback) {
        var options = new HtmlToPdfOptions { AllowSystemFontFallback = fallback };
        options.ResourcePolicy.AllowSystemFontEmbedding = systemFonts;
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<math><mtext>MATH-POLICY</mtext></math>"), options);

        Assert.Empty(result.RenderResult!.Document.Fonts.Faces);
        Assert.DoesNotContain(result.RenderResult.Diagnostics, item => item.Code == "InstalledFontResolved");
        Assert.Contains("MATH-POLICY", OfficeIMO.Pdf.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlMathMlPdf_PortablePolicyRetainsExplicitInMemoryMathematicalFace() {
        var options = new HtmlToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("math", ManagedTextShapingTestAssets.CreateFont('x'));
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<math><mtext>x</mtext></math>"), options);

        Assert.True(result.RenderResult!.Document.Fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out _));
        Assert.DoesNotContain(result.RenderResult.Diagnostics, item => item.Code == "InstalledFontResolved");
        Assert.Equal("x", OfficeIMO.Pdf.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText().Trim());
    }

    [Fact]
    public void HtmlMathMlPdf_BalancedInstalledMathLoadingPreservesSourceByteLimit() {
        if (!HasInstalledMathematicalFace()) return;
        var options = new HtmlToPdfOptions { MaxResourceBytes = 1 };
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<math><mtext>x</mtext></math>"), options);

        Assert.Empty(result.RenderResult!.Document.Fonts.Faces);
        Assert.Contains(result.RenderResult.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded);
        Assert.Equal("x", OfficeIMO.Pdf.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText().Trim());
    }

    [Fact]
    public void HtmlMathMlPdf_SvgForeignObjectsShareTheParentInstalledFontBudget() {
        HtmlRenderDocument probe = HtmlRenderEngine.Render(HtmlConversionDocument.Parse("<math><mtext>x</mtext></math>"));
        if (!probe.Fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out _)) return;
        HtmlDiagnostic loaded = Assert.Single(probe.Diagnostics, item => item.Code == "InstalledFontResolved");
        long decodedBytes = long.Parse(loaded.Detail!.Split('=')[2], System.Globalization.CultureInfo.InvariantCulture);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='40'>"
            + "<foreignObject width='200' height='40'><div xmlns='http://www.w3.org/1999/xhtml' "
            + "style='font-family:math'>NESTED</div></foreignObject></svg>";
        string image = "<img src='data:image/svg+xml;base64,"
            + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(svg)) + "'>";
        var options = new HtmlToPdfOptions {
            MaxResourceBytes = decodedBytes,
            MaxTotalResourceBytes = decodedBytes + System.Text.Encoding.UTF8.GetByteCount(svg) * 2L + 1L
        };
        HtmlPdfRenderResult result = HtmlPdfRenderedConverter.Convert(
            HtmlConversionDocument.Parse("<math><mtext>PARENT</mtext></math>" + image + image), options);
        Assert.Single(result.RenderResult!.Diagnostics, item => item.Code == "InstalledFontResolved");
        Assert.Contains(result.RenderResult.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded);
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(result.Document.ToBytes()).ExtractText();
        Assert.Contains("PARENT", text, StringComparison.Ordinal);
        Assert.Equal(2, text.Split(new[] { "NESTED" }, StringSplitOptions.None).Length - 1);
    }

    private static bool HasInstalledMathematicalFace() => HtmlRenderEngine.Render(
        HtmlConversionDocument.Parse("<math><mtext>x</mtext></math>"))
        .Fonts.TryResolveFaceForText("x", "math", OfficeFontStyle.Regular, out _);
}
