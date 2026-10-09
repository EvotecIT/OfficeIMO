using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void StaticPngDecodeHonorsDefaultCanvasLimitsAndCancellation() {
        byte[] fallback = OfficePngWriter.Encode(new OfficeRasterImage(2, 1, OfficeColor.Red));
        byte[] frame = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Lime));
        byte[] apng = CreateSingleFrameApngAfterFallback(fallback, frame);
        var options = new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 };
        Assert.False(OfficeRasterImageDecoder.TryDecodePngDefault(apng, options, out var rejected));
        Assert.Null(rejected);
        options.MaximumDecodedPixels = 2;
        Assert.True(OfficeRasterImageDecoder.TryDecodePngDefault(apng, options, out var image));
        Assert.Equal(2, image!.Width); Assert.Equal(OfficeColor.Red, image.GetPixel(1, 0));
        options.MaximumEncodedBytes = apng.Length - 1;
        Assert.False(OfficeRasterImageDecoder.TryDecodePngDefault(apng, options, out _));
        options.MaximumEncodedBytes = apng.Length;
        options.CancellationToken = new CancellationToken(canceled: true);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecodePngDefault(apng, options, out _));
    }
}
