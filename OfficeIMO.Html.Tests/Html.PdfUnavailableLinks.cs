using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfUnavailableLinkTests {
    [Theory]
    [InlineData("<a href='terms.html'>Terms</a>")]
    [InlineData("<a href='terms.html'><img alt='Terms' width='20' height='20' src='data:image/png;base64,IMAGE'></a>")]
    [InlineData("<svg width='80' height='40'><a href='terms.html'><rect width='80' height='40' fill='green'/></a></svg>")]
    public void RelativeLinksWithoutABaseDoNotPreventPdfOutput(string body) {
        body = body.Replace("IMAGE", Convert.ToBase64String(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red))));
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p>Body <a href='https://example.test/absolute'>Absolute</a></p>" + body);
        var result = source.ConvertToPdfWithEvidence();

        byte[] bytes = result.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(1, info.PageCount);
        Assert.DoesNotContain("terms.html", info.LinkUris);
        Assert.Contains("https://example.test/absolute", info.LinkUris);
        Assert.Contains("Body", PdfReadDocument.Open(bytes).ExtractText());
        if (body.Contains("<img", StringComparison.Ordinal)) {
            Assert.NotEmpty(PdfReadDocument.Open(bytes).Pages[0].GetImages());
        }
        Assert.Contains(result.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable
            && warning.LossKind == OfficeConversionLossKind.Omission);

        byte[] resolved = source.RenderToPdfBytes(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlPdfExportOptions { BaseUri = new Uri("https://example.test/docs/") }));
        Assert.Contains("https://example.test/docs/terms.html", PdfInspector.Inspect(resolved).LinkUris);
    }

    [Fact]
    public void LinksToUnselectedPagesAreOmittedWhileSelectedTargetsRemainNavigable() {
        const string html = "<style>@page{size:300px 200px;margin:8px}</style>"
            + "<p id='first'><a href='#later'>Later</a> <a href='#first'>First</a></p>"
            + "<h1 id='later' style='break-before:page'>Later page</h1>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderRequest request = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf);
        byte[] complete = source.RenderToPdfBytes(request);
        Assert.Equal(2, PdfInspector.Inspect(complete).PageCount);
        Assert.Contains("html-fragment:later", PdfInspector.Inspect(complete).LinkDestinationNames);

        var result = source.RenderToPdfResult(request.WithPageSet(HtmlRenderPageSet.Page(0)));
        byte[] selected = result.Output.ToBytes();
        PdfDocumentInfo info = PdfInspector.Inspect(selected);

        Assert.Equal(1, info.PageCount);
        Assert.Contains("html-fragment:first", info.LinkDestinationNames);
        Assert.DoesNotContain("html-fragment:later", info.LinkDestinationNames);
        Assert.Contains(result.Output.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.HyperlinkTargetUnavailable
            && warning.LossKind == OfficeConversionLossKind.Omission);
    }
}
