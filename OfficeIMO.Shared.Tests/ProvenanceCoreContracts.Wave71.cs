using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData(32773)]
    [InlineData(8)]
    public void TiffStrictRemovalRejectsMalformedSupportedCompressedPixelPayload(int compression) {
        byte[] malformedPayload = compression == 32773 ? new byte[] { 0xFF } : new byte[] { 0x00 };
        byte[] tiff = CreateCompressedTiff(CreateManifestStore(), compression, malformedPayload);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tiff");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void TiffStrictRemovalAcceptsCompletePackBitsPixelPayload() {
        byte[] tiff = CreateCompressedTiff(
            CreateManifestStore(),
            compression: 32773,
            pixelPayload: new byte[] { 2, 0, 0, 0 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tiff");

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    private static byte[] CreateCompressedTiff(byte[] manifest, int compression, byte[] pixelPayload) {
        const int ifdOffset = 8;
        const int entryCount = 10;
        const int bitsOffset = ifdOffset + 2 + entryCount * 12 + 4;
        int manifestOffset = bitsOffset + 6;
        int pixelOffset = manifestOffset + manifest.Length;
        byte[] result = new byte[pixelOffset + pixelPayload.Length];
        result[0] = result[1] = (byte)'I';
        result[2] = 42;
        result[4] = ifdOffset;
        result[ifdOffset] = entryCount;
        WriteLittleEndianEntry(result, 10, 256, 4, 1, 1);
        WriteLittleEndianEntry(result, 22, 257, 4, 1, 1);
        WriteLittleEndianEntry(result, 34, 258, 3, 3, bitsOffset);
        WriteLittleEndianEntry(result, 46, 259, 3, 1, compression);
        WriteLittleEndianEntry(result, 58, 262, 3, 1, 2);
        WriteLittleEndianEntry(result, 70, 273, 4, 1, pixelOffset);
        WriteLittleEndianEntry(result, 82, 277, 3, 1, 3);
        WriteLittleEndianEntry(result, 94, 278, 4, 1, 1);
        WriteLittleEndianEntry(result, 106, 279, 4, 1, pixelPayload.Length);
        WriteLittleEndianEntry(result, 118, 0xCD41, 7, manifest.Length, manifestOffset);
        result[bitsOffset] = result[bitsOffset + 2] = result[bitsOffset + 4] = 8;
        manifest.CopyTo(result, manifestOffset);
        pixelPayload.CopyTo(result, pixelOffset);
        return result;
    }
}
