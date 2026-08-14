using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void DuplicateNonCarrierTiffTagsInvalidateTheIfd() {
        byte[] xmp = CreateXmpPacket();
        const int payloadOffset = 50;
        byte[] tiff = new byte[payloadOffset + xmp.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 3;
        WriteLittleEndianEntry(tiff, 10, 256, 3, 1, 1);
        WriteLittleEndianEntry(tiff, 22, 256, 3, 1, 1);
        WriteLittleEndianEntry(tiff, 34, 700, 1, xmp.Length, payloadOffset);
        Buffer.BlockCopy(xmp, 0, tiff, payloadOffset, xmp.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void ReservedJumbfDescriptionToggleBitInvalidatesTheManifest() {
        byte[] manifest = CreateManifestStore();
        manifest[32] |= 0x10;
        byte[] png = CreatePngWithC2paManifest(manifest);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }
}
