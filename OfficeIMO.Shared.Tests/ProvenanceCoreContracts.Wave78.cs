using OfficeIMO.Provenance;
using System.Threading;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
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
