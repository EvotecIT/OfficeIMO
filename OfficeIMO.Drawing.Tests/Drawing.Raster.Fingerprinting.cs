using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterFingerprintingTests {
    [Fact]
    public void DifferenceHashUsesHorizontalComparisonsInRowMajorBitOrder() {
        var image = new OfficeRasterImage(9, 8);
        for (int y = 0; y < 8; y++) {
            for (int x = 0; x < 9; x++) {
                byte value = (byte)((x + y) % 2 == 0 ? 220 : 20);
                image.SetPixel(x, y, OfficeColor.FromRgb(value, value, value));
            }
        }
        byte[] original = image.GetPixels();

        Assert.Equal(0xaa55aa55aa55aa55UL, OfficeRasterFingerprinting.DifferenceHash(image));
        Assert.Equal(original, image.GetPixels());
    }

    [Fact]
    public void DifferenceHashUsesLuminanceRatherThanOnlyTheRedChannel() {
        var image = new OfficeRasterImage(9, 8);
        for (int y = 0; y < 8; y++) {
            for (int x = 0; x < 9; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgb(100, (byte)(250 - x * 20), 100));
            }
        }
        Assert.Equal(ulong.MaxValue, OfficeRasterFingerprinting.DifferenceHash(image));
        Assert.Equal(0UL, OfficeRasterFingerprinting.DifferenceHash(new OfficeRasterImage(1, 1, OfficeColor.Red)));
    }

    [Fact]
    public void DifferenceHashObservesCancellation() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterFingerprinting.DifferenceHash(
            new OfficeRasterImage(1, 1), cancellation.Token));
    }
}
