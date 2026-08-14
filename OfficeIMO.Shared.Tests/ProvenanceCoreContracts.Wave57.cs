using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void WebpWithConflictingSimpleImagePayloadsIsPreserved() {
        byte[] webp = CreateWebp(
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("VP8L", new byte[] { 0x2f, 1, 2, 3, 4 }),
            CreateRiffChunk("C2PA", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }
}
