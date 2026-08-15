using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void TrailingUnmatchedStructuredManifestMakesRemovalAmbiguous() {
        string validBlock = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] input = Encoding.UTF8.GetBytes(validBlock + "-----BEGIN C2PA MANIFEST-----\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void TiffRejectsCyclicIfdGraphs() {
        byte[] manifest = CreateManifestStore();
        const int ifdOffset = 8;
        const int payloadOffset = 26;
        byte[] tiff = new byte[payloadOffset + manifest.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = ifdOffset;
        tiff[ifdOffset] = 1;
        WriteLittleEndianEntry(tiff, ifdOffset + 2, 0xCD41, 7, manifest.Length, payloadOffset);
        BitConverter.GetBytes(ifdOffset).CopyTo(tiff, ifdOffset + 14);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(tiff, "fixture.tif"));
    }

    [Fact]
    public void SecondaryIfdC2paTagsInvalidateThePrimaryCarrier() {
        byte[] manifest = CreateManifestStore();
        const int primaryIfdOffset = 8;
        const int subIfdOffset = 38;
        const int payloadOffset = 56;
        byte[] tiff = new byte[payloadOffset + (manifest.Length * 2)];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = primaryIfdOffset;
        tiff[primaryIfdOffset] = 2;
        WriteLittleEndianEntry(tiff, primaryIfdOffset + 2, 330, 4, 1, subIfdOffset);
        WriteLittleEndianEntry(tiff, primaryIfdOffset + 14, 0xCD41, 7, manifest.Length, payloadOffset);
        tiff[subIfdOffset] = 1;
        WriteLittleEndianEntry(tiff, subIfdOffset + 2, 0xCD41, 7, manifest.Length, payloadOffset + manifest.Length);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset + manifest.Length, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }
}
