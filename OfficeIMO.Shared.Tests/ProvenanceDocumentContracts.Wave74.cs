using System.IO.Compression;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ResponsiveScreenMediaRemainsInProvenanceScope() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style media='(max-width: 600px)'>.box{background-image:url('" +
            dataUri + "')}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void BackgroundImageLonghandOverridesEarlierShorthandCarrier() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.box{background:url('" + dataUri +
            "');background-image:none}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void OdfPreservesUnownedSignatureLikeResources() {
        const string resourcePath = "META-INF/audit-signatures.xml";
        byte[] package = CreateZipPackage(
            "odt",
            resourcePath,
            CreatePngWithManifest(CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains(archive.Entries, entry => entry.FullName == resourcePath);
        Assert.Empty(result.After.Evidence);
    }
}
