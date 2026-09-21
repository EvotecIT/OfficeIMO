using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    private static string Wave61DataUri() => "data:image/png;base64," +
        Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));




    [Theory]
    [InlineData("1.x")]
    [InlineData("1.e2x")]
    public void InvalidDensityDescriptorsDoNotSelectSrcsetCandidates(string descriptor) {
        string html = "<img srcset='" + Wave61DataUri() + " " + descriptor + "'>";
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        Assert.Empty(result.Before.Evidence);
    }

    [Fact]
    public void QuotedBracesDoNotBreakFollowingCssDeclarationDiscovery() {
        string html = "<style>.box{content:'{';background-image:url('" + Wave61DataUri() +
            "')}</style><div class='box'></div>";
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
    }


    [Fact]
    public void UnpaddedBase64ImageDataUrisAreSupported() {
        string payload = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore())).TrimEnd('=');
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove("<img src='data:image/png;base64," + payload + "'>");
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }


    [Fact]
    public void XlsbOwnershipRejectsBackslashRelationshipTargets() {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: false);
        const string relationships =
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>" +
            "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl\\workbook.bin'/>" +
            "</Relationships>";
        package = ReplaceWave38Entry(package, "_rels/.rels", relationships);

        Assert.Throws<InvalidDataException>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }
}
