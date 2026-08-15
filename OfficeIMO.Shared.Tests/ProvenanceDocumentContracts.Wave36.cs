using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightDoesNotTreatSlashInUnquotedAttributeAsSelfClosing() {
        string html = "<html><body><svg><foreignObject x=a/><![CDATA[x>" +
            string.Concat(Enumerable.Repeat("<div></div>", 64)) +
            "]]></foreignObject></svg></body></html>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 16 }));
    }
}
