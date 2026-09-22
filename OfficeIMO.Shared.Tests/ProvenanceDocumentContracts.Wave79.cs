using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightAcceptsBalancedOrdinaryNestingWithinEntryLimit() {
        string html = "<html><body>" + string.Concat(Enumerable.Repeat("<div>", 80)) +
            "content" + string.Concat(Enumerable.Repeat("</div>", 80)) + "</body></html>";
        var limits = new OfficeProvenanceOptions { MaxContainerEntries = 100 };

        Assert.Empty(HtmlProvenance.Inspect(html, limits).Evidence);
        Assert.False(HtmlProvenance.Remove(html, new OfficeProvenanceRemovalOptions {
            Limits = { MaxContainerEntries = 100 }
        }).WasChanged);
    }

    [Fact]
    public void HtmlProvenancePreservesSvgDataUriWithMalformedCharset() {
        const string html = "<html><body><img src='data:image/svg+xml;charset=\"unterminated,%3Csvg%3E%3C/svg%3E'></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult removal = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("could not be decoded", StringComparison.Ordinal));
        Assert.False(removal.WasChanged);
    }
}
