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

    [Theory]
    [InlineData(0)]
    [InlineData(2 * 1024 * 1024)]
    public void OdfPreservesUnownedSignatureLikeResources(int paddingBytes) {
        const string resourcePath = "META-INF/audit-signatures.xml";
        byte[] resource = paddingBytes == 0
            ? CreatePngWithManifest(CreateManifestStore())
            : Join(
                new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
                CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
                CreatePngChunk("tEXt", new byte[paddingBytes]),
                CreatePngChunk("IDAT", new byte[] { 0x78, 0x9C, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01 }),
                CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateZipPackage(
            "odt",
            resourcePath,
            CreatePngWithManifest(CreateManifestStore()),
            signatureContent: resource);

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.False(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains(archive.Entries, entry => entry.FullName == resourcePath);
        Assert.Empty(result.After.Evidence);
    }
}
