using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData("# ", "")]
    [InlineData("// ", "")]
    [InlineData("-- ", "")]
    [InlineData("; ", "")]
    [InlineData("% ", "")]
    [InlineData("' ", "")]
    [InlineData("REM ", "")]
    [InlineData(":: ", "")]
    [InlineData("/* ", " */")]
    [InlineData("<!-- ", " -->")]
    [InlineData("<# ", " #>")]
    public void C2pa24SingleLineCommentReferencesAreInspectedAndRemoved(string prefix, string suffix) {
        string carrier = prefix + "-----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa " +
            "-----END C2PA MANIFEST-----" + suffix + "\r\n";
        byte[] input = Encoding.UTF8.GetBytes("before\r\n" + carrier + "after\r\n");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(input, "fixture.py");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.py");

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paExternalManifest, evidence.Carrier);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal("https://example.test/manifest.c2pa", evidence.Value);
        Assert.Equal("before\r\nafter\r\n", Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void C2pa24SingleLineEmbeddedManifestIsInspectedAndRemoved() {
        string encoded = Convert.ToBase64String(CreateManifestStore());
        byte[] input = Encoding.UTF8.GetBytes(
            "# -----BEGIN C2PA MANIFEST----- data:application/c2pa;base64," + encoded +
            " -----END C2PA MANIFEST-----\nbody\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.py");

        OfficeProvenanceEvidence evidence = Assert.Single(result.Before.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, evidence.Carrier);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal("body\n", Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void BareSameLineDelimitersAreNotTreatedAsC2pa24CommentCarriers() {
        byte[] input = Encoding.UTF8.GetBytes(
            "-----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa -----END C2PA MANIFEST-----\nbody\n");

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(input, "fixture.txt");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void DuplicateSingleLineReferencesRemainFailClosed() {
        const string block = "# -----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa -----END C2PA MANIFEST-----\n";
        byte[] input = Encoding.UTF8.GetBytes(block + block);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.py");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Theory]
    [InlineData("This says -----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa -----END C2PA MANIFEST----- as prose.")]
    [InlineData("/* -----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa -----END C2PA MANIFEST----- wrong")]
    [InlineData("# -----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa -----END C2PA MANIFEST----- trailing")]
    public void SingleLineLookalikesAreNotRemoved(string line) {
        byte[] input = Encoding.UTF8.GetBytes(line + "\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void RecognizedCommentWithMissingEndDelimiterIsReportedButPreserved() {
        byte[] input = Encoding.UTF8.GetBytes(
            "# -----BEGIN C2PA MANIFEST----- https://example.test/manifest.c2pa\nbody\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.py");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }
}
