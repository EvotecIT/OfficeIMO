using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
#if SHARED_PERFORMANCE_EVIDENCE
    [Trait("Category", "Performance")]
    [Trait("Category", "ResourcePerformanceEvidence")]
#endif
    public void InlineTextBeginDelimitersAreIgnoredOnOneLongLine() {
#if SHARED_PERFORMANCE_EVIDENCE
        const int delimiters = 8192;
#else
        const int delimiters = 4;
#endif
        byte[] input = Encoding.UTF8.GetBytes("prefix " +
            string.Concat(Enumerable.Repeat("-----BEGIN C2PA MANIFEST-----x", delimiters)));
#if SHARED_PERFORMANCE_EVIDENCE
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
#endif

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            input, "fixture.txt", new OfficeProvenanceOptions { MaxContainerEntries = delimiters });

        Assert.Empty(report.Evidence);
#if SHARED_PERFORMANCE_EVIDENCE
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5));
#endif
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
