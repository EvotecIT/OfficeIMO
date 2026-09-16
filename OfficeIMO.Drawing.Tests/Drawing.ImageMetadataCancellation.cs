using OfficeIMO.Drawing;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ImageMetadataCancellationTests {
    [Fact]
    public void JpegInspectionChecksCancellationIndependentlyOfMarkerOffsets() {
        byte[] jpeg = {
            0xFF, 0xD8,
            0xFF, 0xFE, 0x00, 0x02,
            0xFF, 0xD9
        };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeImageMetadataInspector.InspectJpeg(
                jpeg,
                new OfficeImageMetadataSnapshot(),
                retainedManagedBytes: 0L,
                cancellationToken: cancellation.Token));
    }
}