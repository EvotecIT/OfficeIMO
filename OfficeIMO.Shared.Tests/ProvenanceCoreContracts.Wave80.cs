using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void CarrierPngChargesItsDecodeToEachInspection() {
        byte[] image = CreatePngWithC2paManifest(CreateManifestStore());
        Assert.True(OfficePngReader.TryGetProvenanceDecodeBudget(image, default, int.MaxValue, out long decodedBytes));

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            image, "image.png", new OfficeProvenanceOptions { MaxExpandedContainerBytes = decodedBytes - 1 }));
        Assert.True(OfficeProvenanceInspector.Inspect(
            image, "image.png", new OfficeProvenanceOptions { MaxExpandedContainerBytes = decodedBytes }).HasC2paManifest);

        // Removal reuses the inspected carrier validity, and the carrier-free output is not decoded.
        var removal = new OfficeProvenanceRemovalOptions();
        removal.Limits.MaxExpandedContainerBytes = decodedBytes;
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(image, "image.png", removal);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void CarrierFreeEmbeddedPngIsNotDecoded() {
        byte[] image = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", CreateValidPngHeader()),
            CreatePngChunk("IDAT", CreateValidPngImageData()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateCompressedZip(("media/plain.png", image));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            package, "package.zip", new OfficeProvenanceOptions { MaxExpandedContainerBytes = image.Length });

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void PngRemovalKeepsACarrierWhoseScanlinesDoNotDecode() {
        // A stored deflate block holding one scanline with the undefined filter type 5.
        byte[] invalidImageData = { 0x78, 0x01, 0x01, 0x04, 0x00, 0xFB, 0xFF, 0x05, 0x00, 0x00, 0x00, 0x00, 0x18, 0x00, 0x06 };
        byte[] image = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", CreateValidPngHeader()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", invalidImageData),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        Assert.False(Assert.Single(OfficeProvenanceInspector.Inspect(image, "image.png").Evidence).IsStructurallyValid);
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(image, "image.png");

        Assert.False(result.WasChanged);
        Assert.Equal(image, result.ToArray());
    }

    [Fact]
    public void HtmlEmbeddedPngDecodeSharesTheDocumentBudget() {
        byte[] image = CreatePngWithC2paManifest(CreateManifestStore());
        Assert.True(OfficePngReader.TryGetProvenanceDecodeBudget(image, default, int.MaxValue, out long decodedBytes));
        string html = "<html><body><img src='data:image/png;base64," + Convert.ToBase64String(image) + "'></body></html>";
        long required = image.Length + decodedBytes;

        Assert.True(HtmlProvenance.Inspect(html, new OfficeProvenanceOptions { MaxExpandedContainerBytes = required }).HasC2paManifest);
        // The nested image's limit breach propagates instead of being reported as a malformed image.
        Assert.Throws<InvalidDataException>(() =>
            HtmlProvenance.Inspect(html, new OfficeProvenanceOptions { MaxExpandedContainerBytes = required - 1 }));
    }
}
