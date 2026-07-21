using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlRtfStructuredResult {
    [Fact]
    public void HtmlRtf_UsesSharedResultAndLossDiagnosticsWithoutReparsingSharedDocument() {
        HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse("<p><unknown>Value</unknown></p>");
        var options = new HtmlToRtfOptions { PreserveUnknownTagsAsText = true };

        HtmlToRtfResult import = source.ToRtfDocumentResult(options);
        import.Value.AddParagraph().AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        RtfToHtmlResult export = import.Value.ToHtmlResult();

        Assert.True(import.Succeeded);
        Assert.True(export.Succeeded);
        Assert.Contains("Value", export.Value, StringComparison.Ordinal);
        Assert.Equal(import.Report.Diagnostics.Count, import.RtfDiagnostics.Count);
        Assert.Contains(export.Report.Diagnostics, diagnostic => diagnostic.LossKind == HtmlConversionLossKind.Omission);
    }

    [Fact]
    public void HtmlRtf_UsesSharedDocumentBaseUriForRelativeHyperlinks() {
        HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<p><a href='docs/start.html'>Guide</a></p>",
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri("https://example.test/root/page.html")
            });

        HtmlToRtfResult import = source.ToRtfDocumentResult();

        RtfRun run = Assert.Single(Assert.Single(import.Value.Paragraphs).Runs);
        Assert.Equal(new Uri("https://example.test/root/docs/start.html"), run.Hyperlink);
    }
}
