using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("right-middle", false)]
    [InlineData("top-right-corner", false)]
    [InlineData("bottom-right-corner", false)]
    [InlineData("right-middle", true)]
    public async Task HtmlPageMarginPrintFit_StationeryDoesNotShrinkOrdinaryBody(string position, bool running) {
        string html = PageMarginPrintFitHtml(position, running, wideBody: false);
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlToPdfOptions options = PageMarginPrintFitOptions();
        HtmlRenderResult automatic = await HtmlRenderEngine.ExecuteAsync(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        options.AutoFitWidePrintContent = false;
        HtmlRenderResult reference = await HtmlRenderEngine.ExecuteAsync(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));

        HtmlRenderPage page = Assert.Single(automatic.Document.Pages);
        Assert.Equal(400D, page.Width, 6);
        Assert.Equal(400D, page.Height, 6);
        Assert.Equal(Assert.Single(reference.ExportImages()).Bytes, Assert.Single(automatic.ExportImages()).Bytes);
        Assert.Equal("BODY", automatic.Document.Text.Trim());
        automatic.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HtmlPageMarginPrintFit_StationeryPreservesRealBodyOverflowFitting(bool running) {
        HtmlToPdfOptions options = PageMarginPrintFitOptions();
        HtmlRenderResult ordinary = await HtmlRenderEngine.ExecuteAsync(
            HtmlConversionDocument.Parse(PageMarginPrintFitHtml(null, running: false, wideBody: true)),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        HtmlRenderResult decorated = await HtmlRenderEngine.ExecuteAsync(
            HtmlConversionDocument.Parse(PageMarginPrintFitHtml("right-middle", running, wideBody: true)),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        HtmlRenderPage expected = Assert.Single(ordinary.Document.Pages);
        HtmlRenderPage actual = Assert.Single(decorated.Document.Pages);
        Assert.True(expected.Width > 400D);
        Assert.Equal(expected.Width, actual.Width, 6);
        Assert.Equal(expected.Height, actual.Height, 6);
        Assert.Equal("BODY", decorated.Document.Text.Trim());
        decorated.RequireNoLoss();
    }

    private static string PageMarginPrintFitHtml(string? position, bool running, bool wideBody) {
        string family = ManagedTextShapingTestAssets.FamilyName;
        string box = position == null ? string.Empty : "@" + position + "{content:"
            + (running ? "element(folio)" : "'F'") + ";width:40px;height:40px;background:blue}";
        string snapshot = running ? "<div style='position:running(folio);width:20px;height:20px'>F</div>" : string.Empty;
        return "<style>html,body,div{margin:0;padding:0;font:20px/20px '" + family
            + "'}@page{size:400px 400px;margin:60px 80px 40px 100px;font:20px/20px '" + family
            + "';" + box + "}</style>" + snapshot + "<div" + (wideBody ? " style='width:440px'" : "") + ">BODY</div>";
    }

    private static HtmlToPdfOptions PageMarginPrintFitOptions() {
        string family = ManagedTextShapingTestAssets.FamilyName;
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithScalarAdvances(
            Enumerable.Range(32, 95).ToDictionary(scalar => scalar, _ => (ushort)500));
        var options = new HtmlToPdfOptions {
            PageSize = new OfficePageSize(400D / 96D, 400D / 96D), Margins = HtmlRenderMargins.All(0D),
            ViewportWidth = 400D, ViewportHeight = 400D, AllowSystemFontFallback = false,
            DefaultFontFamily = family, FontFamily = new PdfEmbeddedFontFamily(family, font),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser
        };
        options.Fonts.Add(family, font);
        return options;
    }
}
