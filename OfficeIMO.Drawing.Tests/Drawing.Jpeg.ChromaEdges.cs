using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegChromaEdgeTests {
    [Theory]
    [InlineData(OfficeJpegSubsampling.Y420, false)]
    [InlineData(OfficeJpegSubsampling.Y420, true)]
    [InlineData(OfficeJpegSubsampling.Y422, false)]
    [InlineData(OfficeJpegSubsampling.Y422, true)]
    public void SmoothChromaClampsAtVisibleEvenDimensionCorner(OfficeJpegSubsampling subsampling, bool progressive) {
        const int size = 10;
        var source = new OfficeRasterImage(size, size);
        for (int y = 0; y < size; y++) {
            for (int x = 0; x < size; x++) {
                source.SetPixel(x, y, OfficeColor.FromRgb(
                    (byte)(16 + x * 150 / (size - 1)), (byte)(24 + y * 96 / (size - 1)), (byte)(32 + (x + y) * 40 / (size - 1))));
            }
        }
        byte[] encoded = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 85, Subsampling = subsampling, Progressive = progressive
        });

        OfficeRasterImage nearest = OfficeJpegCodec.Decode(encoded);
        OfficeRasterImage smooth = OfficeJpegCodec.Decode(encoded, new OfficeJpegDecodeOptions(highQualityChroma: true));

        // For even dimensions, this corner's smooth filter extends beyond the
        // visible component. Edge replication must use the last real sample,
        // so block padding cannot change its color.
        Assert.Equal(nearest.GetPixel(size - 1, size - 1), smooth.GetPixel(size - 1, size - 1));
    }
}
