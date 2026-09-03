using OfficeIMO.Provenance;
using System.Threading;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void ZipInspectionReportsExpandedManifestBytes() {
        byte[] manifest = CreateManifestStore();
        byte[] package = CreateZip(("META-INF/content_credential.c2pa", manifest));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(package, "fixture.zip");

        Assert.Equal(manifest.LongLength, report.ExpandedInspectionBytes);
    }

    [Fact]
    public void StructuredTextDetectionObservesCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeProvenanceText.HasStructuredDelimiter(new byte[16 * 1024], cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() =>
            OfficeProvenanceText.HasUnstructuredWrapperPrefix(
                new byte[16 * 1024],
                maximumContainerEntries: 16,
                cancellation.Token));
    }
}
