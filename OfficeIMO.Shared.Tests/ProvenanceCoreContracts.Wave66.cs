using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void C2paManifestByteLimitIsNotReportedAsMalformedEvidence() {
        byte[] manifest = CreateManifestStore();

        Assert.Throws<InvalidDataException>(() => OfficeC2paManifestStore.IsValid(
            manifest, 0, manifest.Length, manifest.Length - 1, int.MaxValue, out _));
    }

    [Fact]
    public void C2paManifestEntryLimitIsNotReportedAsMalformedEvidence() {
        byte[] manifest = CreateManifestStore();

        Assert.Throws<InvalidDataException>(() => OfficeC2paManifestStore.IsValid(
            manifest, 0, manifest.Length, manifest.Length, 1, out _));
    }

    [Fact]
    public void IncompleteJpegPayloadPreservesOtherwiseValidC2paCarrier() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            CreateMinimalJpegFrame(),
            CreateMinimalJpegScan(),
            new byte[] { 0, 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }
}
