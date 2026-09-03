using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData(
        "<link rel=\"c2pa-manifest\" href=\"claim.c2pa\">",
        "claim.c2pa")]
    [InlineData(
        "<base href=\"https://example.test/assets/\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\">",
        "https://example.test/assets/claim.c2pa")]
    [InlineData(
        "<title>parser-recovered metadata</title><link rel=\"c2pa-manifest\" href=\"claim.c2pa\">",
        "claim.c2pa")]
    [InlineData(
        "<head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>",
        "claim.c2pa")]
    public void CoreHtmlMatchesParserRecoveredAfterHeadExternalManifest(
        string afterHead,
        string expectedReference) {
        string html = "<!doctype html><html><head></head>" + afterHead + "<body>body</body></html>";

        OfficeProvenanceReport core = OfficeProvenanceInspector.Inspect(
            Encoding.UTF8.GetBytes(html),
            "fixture.html");
        OfficeProvenanceReport owner = HtmlProvenance.Inspect(html);

        OfficeProvenanceEvidence coreEvidence = Assert.Single(core.Evidence);
        OfficeProvenanceEvidence ownerEvidence = Assert.Single(owner.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paExternalManifest, coreEvidence.Carrier);
        Assert.Equal(ownerEvidence.Value, coreEvidence.Value);
        Assert.Equal(expectedReference, core.GetExternalManifestReference(coreEvidence));
        Assert.Equal(expectedReference, owner.GetExternalManifestReference(ownerEvidence));
    }

    [Fact]
    public void CoreHtmlMatchesParserRecoveredAfterHeadEmbeddedManifest() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<!doctype html><html><head></head><script type=\"application/c2pa\">" +
            manifest + "</script><body>body</body></html>";

        OfficeProvenanceReport core = OfficeProvenanceInspector.Inspect(
            Encoding.UTF8.GetBytes(html),
            "fixture.html");
        OfficeProvenanceReport owner = HtmlProvenance.Inspect(html);

        OfficeProvenanceEvidence coreEvidence = Assert.Single(core.Evidence);
        OfficeProvenanceEvidence ownerEvidence = Assert.Single(owner.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, coreEvidence.Carrier);
        Assert.True(coreEvidence.IsStructurallyValid);
        Assert.Equal(ownerEvidence.IsStructurallyValid, coreEvidence.IsStructurallyValid);
    }

    [Theory]
    [InlineData("<div>body begins</div>")]
    [InlineData("<noscript>body begins</noscript>")]
    public void CoreHtmlDoesNotRecoverHeadMetadataAfterBodyHasStarted(string bodyStart) {
        string html = "<!doctype html><html><head></head>" + bodyStart +
            "<link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></html>";

        OfficeProvenanceReport core = OfficeProvenanceInspector.Inspect(
            Encoding.UTF8.GetBytes(html),
            "fixture.html");
        OfficeProvenanceReport owner = HtmlProvenance.Inspect(html);

        Assert.Empty(core.Evidence);
        Assert.Empty(owner.Evidence);
    }
}
