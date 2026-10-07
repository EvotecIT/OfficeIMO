using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegColorConversionTests {
    [Theory]
    [InlineData(OfficeJpegSubsampling.Y444, false)]
    [InlineData(OfficeJpegSubsampling.Y422, false)]
    [InlineData(OfficeJpegSubsampling.Y420, false)]
    [InlineData(OfficeJpegSubsampling.Y444, true)]
    [InlineData(OfficeJpegSubsampling.Y422, true)]
    [InlineData(OfficeJpegSubsampling.Y420, true)]
    public void BlockAlignedColorTilesDecodeIdenticallyAcrossSmallAndLargeImages(OfficeJpegSubsampling sampling, bool progressive) {
        OfficeRasterImage small = CreateTiles(32, 32);
        OfficeRasterImage large = CreateTiles(512, 513);
        var options = new OfficeRasterEncodingOptions { WriteResolutionMetadata = false,
            Jpeg = new OfficeJpegEncodeOptions { Quality = 85, Subsampling = sampling, Progressive = progressive } };
        byte[] reference = OfficeJpegCodec.Decode(OfficeRasterImageEncoder.Encode(small, OfficeImageExportFormat.Jpeg, options)).GetPixels();
        byte[] encoded = OfficeRasterImageEncoder.Encode(large, OfficeImageExportFormat.Jpeg, options);
        byte[] actual = OfficeJpegCodec.Decode(encoded).GetPixels();
        var expected = new byte[large.Width * large.Height * 4];
        for (int y = 0; y < large.Height; y++) {
            for (int x = 0; x < large.Width; x++) {
                Buffer.BlockCopy(reference, ((y % 32) * 32 + x % 32) * 4, expected, (y * large.Width + x) * 4, 4);
            }
        }
        Assert.Equal(expected, actual);
    }

    private static OfficeRasterImage CreateTiles(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        OfficeColor[] colors = { OfficeColor.FromRgb(33, 81, 171), OfficeColor.FromRgb(181, 35, 93),
            OfficeColor.FromRgb(57, 191, 103), OfficeColor.FromRgb(213, 143, 47) };
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) image.SetPixel(x, y, colors[((y / 16) & 1) * 2 + ((x / 16) & 1)]);
        }
        return image;
    }
}
