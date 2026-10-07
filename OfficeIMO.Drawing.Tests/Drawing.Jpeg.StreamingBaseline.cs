using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegStreamingBaselineTests {
    [Theory]
    [InlineData(OfficeJpegSubsampling.Y444, false)]
    [InlineData(OfficeJpegSubsampling.Y422, false)]
    [InlineData(OfficeJpegSubsampling.Y420, false)]
    [InlineData(OfficeJpegSubsampling.Y420, true)]
    public void BaselineBlocksPreservePixelsAcrossOddEdgesAndScanlinePadding(OfficeJpegSubsampling subsampling, bool grayscale) {
        const int width = 19, height = 13, stride = width * 4 + 8;
        var rgba = new byte[height * stride];
        var scanlines = new byte[height * (stride + 1)];
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                int offset = y * stride + x * 4;
                rgba[offset] = (byte)(x * 11);
                rgba[offset + 1] = grayscale ? rgba[offset] : (byte)(y * 17);
                rgba[offset + 2] = grayscale ? rgba[offset] : (byte)((x * 7 + y * 13) & 255);
                rgba[offset + 3] = (byte)(64 + ((x * 9 + y * 5) & 191));
            }
            // The prefix is not RGBA data and must never affect the encoded blocks.
            scanlines[y * (stride + 1)] = 123;
            Buffer.BlockCopy(rgba, y * stride, scanlines, y * (stride + 1) + 1, stride);
        }
        var options = new OfficeJpegEncodeOptions { Quality = 85, Subsampling = subsampling };
        byte[] baseline = OfficeJpegWriter.WriteRgba(width, height, rgba, stride, options);
        Assert.Equal(baseline, OfficeJpegWriter.WriteRgbaScanlines(width, height, scanlines, stride, options));
        options.OptimizeHuffman = true;
        byte[] retained = OfficeJpegWriter.WriteRgba(width, height, rgba, stride, options);
        Assert.Equal(OfficeJpegCodec.Decode(retained).GetPixels(), OfficeJpegCodec.Decode(baseline).GetPixels());
    }
}
