using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("width:1200px")]
    [InlineData("min-width:1200px")]
    [InlineData("width:1000px;padding:100px;box-sizing:content-box")]
    public void HtmlPdf_AutomaticPrintFitIncludesUnpaintedDescendantBoxes(string geometry) {
        string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style>"
            + "<div style='" + geometry + ";height:40px'></div>";
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false }));

        Assert.InRange(result.RenderResult.Document.Pages[0].Width, 1199.99D, 1200.01D);
        (double width, _) = PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetPageSize();
        Assert.Equal(OfficePageSizes.A4.WidthInches * 72D, width, 2);
        Assert.All(result.RenderResult.Document.Pages[0].Visuals.Where(visual => visual.Kind == HtmlRenderVisualKind.Shape),
            visual => Assert.IsType<HtmlRenderShape>(visual));
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitIncludesNormalFlowTablesAndKeepsPhysicalMedia() {
        const string html = """
            <style>@page{size:A4;margin:0}html,body{margin:0}#wide{display:none}
            @media(min-width:900px){#wide{display:block}#narrow{display:none}}</style>
            <p id='narrow'>PhysicalMedia</p><p id='wide'>ExpandedMedia</p>
            <table style='width:1200px;table-layout:fixed;border-collapse:collapse'>
              <colgroup><col style='width:600px'><col style='width:600px'></colgroup>
              <tr><td style='padding:0'>FLOWLEFT</td><td style='padding:0'>FLOWRIGHT</td></tr>
            </table>
            """;
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false
        });
        PdfCore.PdfDocumentReadResult read = PdfCore.PdfDocumentReadResult.Load(bytes);
        PdfCore.PdfLogicalTextBlock left = Assert.Single(read.Pages[0].TextBlocks, block => block.Text.Contains("FLOWLEFT", StringComparison.Ordinal));
        PdfCore.PdfLogicalTextBlock right = Assert.Single(read.Pages[0].TextBlocks, block => block.Text.Contains("FLOWRIGHT", StringComparison.Ordinal));
        Assert.InRange(right.XStart - left.XStart, 297.5D, 297.8D);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("PhysicalMedia", text, StringComparison.Ordinal);
        Assert.DoesNotContain("ExpandedMedia", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("auto")]
    [InlineData("clip")]
    public void HtmlPdf_AutomaticPrintFitExcludesClippedDescendantBoxes(string overflow) {
        string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style>"
            + "<div style='width:300px;overflow:" + overflow + "'><div style='width:1200px;height:40px;background:blue'></div></div>";
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false };
        byte[] fitted = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        options.AutoFitWidePrintContent = false;
        byte[] unscaled = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        Assert.Equal(unscaled, fitted);
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitExcludesShadowPaint() {
        const string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style><div style='width:300px;height:40px;box-shadow:1000px 0 0 red'></div>";
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false };
        byte[] fitted = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        options.AutoFitWidePrintContent = false;
        Assert.Equal(HtmlConversionDocument.Parse(html).ToPdfBytes(options), fitted);
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitIncludesTranslatedDescendantBoxes() {
        const string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style><div style='width:900px;height:40px;transform:translateX(300px)'></div>";
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false }));
        Assert.InRange(result.RenderResult.Document.Pages[0].Width, 1199.99D, 1200.01D);
    }
    [Fact]
    public void HtmlPdf_AutomaticPrintFitReportsNonConvergingPercentageOverflow() {
        const string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style><div style='width:120%;height:40px'></div>";
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false }));
        Assert.Contains(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitExcludesTextShadowPaint() {
        const string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style><p style='text-shadow:1000px 0 red'>Shadow text</p>";
        var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false };
        byte[] fitted = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        options.AutoFitWidePrintContent = false;
        Assert.Equal(HtmlConversionDocument.Parse(html).ToPdfBytes(options), fitted);
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitMeasuresLaterTransformedFragments() {
        const string html = "<style>@page{size:A4;margin:0}html,body{margin:0}</style><div style='width:100px;height:2000px;transform:skewX(45deg);transform-origin:0 0'></div>";
        HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false }));
        Assert.InRange(result.RenderResult.Document.Pages[0].Width, 2099.99D, 2100.01D);
        Assert.DoesNotContain(result.RenderResult.Document.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PrintFitOverflowUnresolved);
    }

    [Fact]
    public void HtmlPdf_AutomaticPrintFitIncludesPositionedOverflow() {
        const string html = "<main><p>Retained paragraph</p><p style='position:absolute;left:10000px;top:20px'>Outside paragraph</p></main>";
        string text = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes()).ExtractText();
        Assert.Contains("Retained paragraph", text, StringComparison.Ordinal);
        Assert.Contains("Outside paragraph", text, StringComparison.Ordinal);
    }

}
