using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ManifestRequiredChildrenMustAppearInSpecificationOrder() {
        byte[] manifest = CreateManifestStore();
        int manifestOffset = 8 + ReadBigEndianLength(manifest, 8);
        int firstChild = manifestOffset + 8 + ReadBigEndianLength(manifest, manifestOffset + 8);
        int firstLength = ReadBigEndianLength(manifest, firstChild);
        int secondChild = firstChild + firstLength;
        int secondLength = ReadBigEndianLength(manifest, secondChild);
        byte[] reordered = new byte[manifest.Length];
        Buffer.BlockCopy(manifest, 0, reordered, 0, firstChild);
        Buffer.BlockCopy(manifest, secondChild, reordered, firstChild, secondLength);
        Buffer.BlockCopy(manifest, firstChild, reordered, firstChild + secondLength, firstLength);
        Buffer.BlockCopy(manifest, secondChild + secondLength, reordered, firstChild + secondLength + firstLength,
            manifest.Length - secondChild - secondLength);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(reordered), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void DuplicateGifC2paApplicationsAreStructurallyInvalid() {
        byte[] application = CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore());
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], application, application, new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }

    private static int ReadBigEndianLength(byte[] data, int offset) =>
        data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3];
}
