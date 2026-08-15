using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void PngWithUndecodableImagePayloadPreservesProvenance() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", CreateValidPngHeader()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", new byte[] { 1, 2, 3 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void OversizedExtensionlessZipEntryCannotBypassEmbeddedAssetLimits() {
        byte[] image = new byte[16 * 1024];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }.CopyTo(image, 0);
        byte[] package = CreateCompressedZip(("media/cover", image));
        long maximumAssetBytes = Math.Max(package.Length, 1024);
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = maximumAssetBytes,
            MaxManifestBytes = Math.Min(512, maximumAssetBytes)
        };

        Assert.True(image.LongLength > options.MaxAssetBytes);
        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(package, "fixture.zip", options));
    }

    private static byte[] CreateValidPngImageData() =>
        new byte[] { 0x78, 0x9C, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01 };
}
