using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void DuplicateKnownPngAncillaryChunksPreserveProvenance() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", CreateValidPngHeader()),
            CreatePngChunk("sRGB", new byte[] { 0 }),
            CreatePngChunk("sRGB", new byte[] { 0 }),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", Array.Empty<byte>()),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void MalformedWebpImagePayloadPreservesProvenance() {
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: false),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2, 3 }),
            CreateRiffChunk("C2PA", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void TiffWithConflictingImageStorageRepresentationsPreservesProvenance() {
        byte[] manifest = CreateManifestStore();
        const int payloadOffset = 98;
        int pixelOffset = payloadOffset + manifest.Length;
        byte[] tiff = new byte[pixelOffset + 1];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 7;
        WriteLittleEndianEntry(tiff, 10, 256, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 22, 257, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 34, 273, 4, 1, pixelOffset);
        WriteLittleEndianEntry(tiff, 46, 279, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 58, 324, 4, 1, pixelOffset);
        WriteLittleEndianEntry(tiff, 70, 325, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 82, 0xCD41, 7, manifest.Length, payloadOffset);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ZipEmbeddedAssetCarrierLimitsPropagate() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", CreateValidPngHeader()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", Array.Empty<byte>()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] package = CreateZip(("media/image.png", png));
        var options = new OfficeProvenanceOptions { MaxCarriers = 1 };

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(package, "fixture.zip", options));
    }

    private static byte[] CreateValidVp8Chunk() => CreateRiffChunk("VP8 ", new byte[] {
        0, 0, 0, 0x9D, 0x01, 0x2A, 1, 0, 1, 0
    });
}
