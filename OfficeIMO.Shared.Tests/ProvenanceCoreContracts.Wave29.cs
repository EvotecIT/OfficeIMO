using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void WebpStrictRemovalPreservesDuplicateC2paChunks() {
        byte[] first = CreateRiffChunk("C2PA", CreateManifestStore());
        byte[] second = CreateRiffChunk("C2PA", CreateManifestStore());
        byte[] webp = CreateWebp(first, second);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }
}
