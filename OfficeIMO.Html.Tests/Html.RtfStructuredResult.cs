using OfficeIMO.Html;
using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlRtfStructuredResult {
    [Fact]
    public void HtmlRtf_UsesSharedResultAndLossDiagnosticsWithoutReparsingSharedDocument() {
        HtmlConversionDocument source = HtmlConversionDocumentBuilder.Build("<p><unknown>Value</unknown></p>");
        var options = new HtmlToRtfOptions { PreserveUnknownTagsAsText = true };

        HtmlToRtfResult import = source.ToRtfDocumentResult(options);
        import.Document.AddParagraph().AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3 });
        RtfToHtmlResult export = import.Document.ToHtmlResult();

        Assert.True(import.Succeeded);
        Assert.True(export.Succeeded);
        Assert.Contains("Value", export.Html, StringComparison.Ordinal);
        Assert.Equal(options.HtmlDiagnostics.Count, options.Diagnostics.Count);
        Assert.Contains(export.Diagnostics, diagnostic => diagnostic.LossKind == HtmlConversionLossKind.Omission);
    }
}
