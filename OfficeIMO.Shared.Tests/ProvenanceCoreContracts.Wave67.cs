using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ManyInlineTextBeginDelimitersDoNotRescanOneLongLine() {
        byte[] input = Encoding.UTF8.GetBytes("prefix " +
            string.Concat(Enumerable.Repeat("-----BEGIN C2PA MANIFEST-----x", 8192)));
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            input, "fixture.txt", new OfficeProvenanceOptions { MaxContainerEntries = 8192 });

        Assert.Empty(report.Evidence);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void OrphanStructuredTextEndDelimiterMakesRemovalAmbiguous() {
        string validBlock = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] input = Encoding.UTF8.GetBytes(validBlock + "-----END C2PA MANIFEST-----\n");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void OversizedStructuredTextManifestBlocksPermissiveRemoval() {
        byte[] manifest = CreateManifestStore();
        byte[] input = Encoding.UTF8.GetBytes(
            "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(manifest) + "\n" +
            "-----END C2PA MANIFEST-----\n");
        var options = new OfficeProvenanceRemovalOptions { RequireStructurallyValidCarrier = false };
        options.Limits.MaxManifestBytes = 8;

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(input, "fixture.txt", options));
    }
}
