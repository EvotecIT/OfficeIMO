using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageCancellationContracts {
    [Fact]
    public void RasterToPngConversionObservesCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeImagePngConverter.TryConvertToPng(
                new byte[] { 1, 2, 3, 4 },
                new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token },
                out _,
                out _));
    }

    [Fact]
    public void OrientationNormalizationObservesCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeImageOrientationNormalizer.TryNormalizeToPng(
                new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 },
                applyEmbeddedOrientation: true,
                cancellationToken: cancellation.Token,
                out _,
                out _));
    }
}
