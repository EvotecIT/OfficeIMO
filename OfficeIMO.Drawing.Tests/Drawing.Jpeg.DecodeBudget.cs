using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegDecodeBudgetTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BaselineDecodeReservesEachSharedAcLookupOnce(bool color) {
        byte[] pixel = color ? new byte[] { 120, 42, 217, 255 } : new byte[] { 128, 128, 128, 255 };
        byte[] jpeg = OfficeJpegWriter.WriteRgba(1, 1, pixel, 4,
            new OfficeJpegEncodeOptions { Quality = 85, Subsampling = OfficeJpegSubsampling.Y444 });
        int components = color ? 3 : 1;
        int acTables = color ? 2 : 1;
        // The tiny frame needs one 8x8 block per component: its byte plane,
        // 64 coefficients, 64 output samples and 64 workspace integers.
        long workingBytes = jpeg.LongLength + 4L + 64L * 1024L + components * 640L + acTables * 1024L;
        long retainedBytes = OfficeRasterGuards.MaximumDecodedBytes - workingBytes;

        Assert.True(OfficeJpegCodec.TryDecode(jpeg, CancellationToken.None,
            retainedManagedBytes: retainedBytes, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(1, image!.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal(OfficeJpegCodec.Decode(jpeg).GetPixels(), image.GetPixels());

        Assert.False(OfficeJpegCodec.TryDecode(jpeg, CancellationToken.None,
            retainedManagedBytes: retainedBytes + 1L, out OfficeRasterImage? rejected));
        Assert.Null(rejected);
    }
}
