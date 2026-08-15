using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void MalformedCompetingJpegC2paCandidateInvalidatesTheValidSequence() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            CreateJpegApp11(new byte[] { 1, 2, 3 }, 0, 3, instance: 2, sequence: 1),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }
}
