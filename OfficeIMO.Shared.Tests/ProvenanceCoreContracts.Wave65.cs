using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void MalformedGifFramePreservesC2pa() {
        byte[] malformedFrame = CreateMinimalGifImage();
        malformedFrame[5] = 0;
        malformedFrame[6] = 0;
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[7],
            CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore()),
            malformedFrame,
            new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(gif, result.ToArray());
    }
}
