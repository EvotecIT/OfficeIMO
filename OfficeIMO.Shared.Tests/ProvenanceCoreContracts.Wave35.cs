using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void AssertionUuidContentRequiresExtendedTypeAndPayload() {
        byte[] manifest = CreateManifestStore();
        int contentType = FindAscii(manifest, "cbor");
        Assert.True(contentType >= 0);
        Encoding.ASCII.GetBytes("uuid").CopyTo(manifest, contentType);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(manifest), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void DuplicateWebpXmpChunksAreAllStructurallyInvalid() {
        byte[] xmp = CreateXmpPacket();
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: true),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("XMP ", xmp),
            CreateRiffChunk("XMP ", xmp));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void UnsortedPrimaryTiffTagsMakeC2paStructurallyInvalid() {
        byte[] manifest = CreateManifestStore();
        const int payloadOffset = 38;
        byte[] tiff = new byte[payloadOffset + manifest.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 2;
        WriteLittleEndianEntry(tiff, 10, 0xCD41, 7, manifest.Length, payloadOffset);
        WriteLittleEndianEntry(tiff, 22, 700, 1, 1, 0);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.False(Assert.Single(result.Before.Evidence, evidence => evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }
}
