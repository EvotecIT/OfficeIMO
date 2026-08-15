using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void JpegStrictRemovalPreservesDuplicateManifestSequences() {
        byte[] manifest = CreateManifestStore();
        byte[] first = CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1);
        byte[] second = CreateJpegApp11(manifest, 0, manifest.Length, instance: 2, sequence: 1);
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, first, second, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void ZipStrictRemovalPreservesDuplicateManifestEntries() {
        byte[] package = CreateZip(
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("content.txt", Encoding.UTF8.GetBytes("keep")));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(package, result.ToArray());
    }
}
