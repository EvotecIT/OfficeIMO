using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_PageFontDiscoverySharesTheLayoutOperationLimit() {
        string html = "<style>" + string.Concat(Enumerable.Range(0, 2000)
            .Select(index => "@page p" + index + "{margin-top:60px}"))
            + "@page{@top-center{content:'ł';font:16px Arial}}body{margin:0}</style><div>Body</div>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var options = new HtmlToPdfOptions { MaxLayoutOperations = 10_000 };

        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlPdfRenderedConverter.Convert(document, options));

        Assert.Equal(HtmlRenderDiagnosticCodes.LayoutOperationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), error.LimitSource);
        Assert.Contains("page-margin font", error.Message, StringComparison.Ordinal);
        options.MaxLayoutOperations = 100_000;
        Assert.Single(HtmlPdfRenderedConverter.Convert(document, options).RenderResult!.Document.Pages);
    }
}
