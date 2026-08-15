using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void UnknownReachableTiffFieldTypeInvalidatesStrictCarrierMutation() {
        byte[] manifest = CreateManifestStore();
        const int payloadOffset = 86;
        int pixelOffset = payloadOffset + manifest.Length;
        byte[] tiff = new byte[pixelOffset + 1];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 6;
        WriteLittleEndianEntry(tiff, 10, 256, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 22, 257, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 34, 273, 4, 1, pixelOffset);
        WriteLittleEndianEntry(tiff, 46, 279, 4, 1, 1);
        WriteLittleEndianEntry(tiff, 58, 0xCD41, 7, manifest.Length, payloadOffset);
        WriteLittleEndianEntry(tiff, 70, 65000, 99, 1, 0);
        manifest.CopyTo(tiff, payloadOffset);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tiff");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }
}
