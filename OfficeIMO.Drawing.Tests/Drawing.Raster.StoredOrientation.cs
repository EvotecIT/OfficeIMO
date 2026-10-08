using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    public void StoredSampleDecodingDefersOrientationWithoutChangingTheDefault(OfficeImageExportFormat format) {
        var source = new OfficeRasterImage(3, 2);
        source.SetPixel(0, 0, OfficeColor.Red);
        source.SetPixel(1, 0, OfficeColor.Lime);
        source.SetPixel(2, 0, OfficeColor.Blue);
        source.SetPixel(0, 1, OfficeColor.White);
        source.SetPixel(1, 1, OfficeColor.Black);
        source.SetPixel(2, 1, OfficeColor.Yellow);
        byte[] encoded = OfficeRasterImageEncoder.Encode(source, format);
        var metadata = OfficeImageMetadata.Read(encoded);
        metadata.SetExifValue(OfficeExifTag.Orientation, (ushort)6);
        byte[] tagged = OfficeImageMetadata.Apply(encoded, metadata);

        Assert.True(OfficeRasterImageDecoder.TryDecode(tagged, out var display));
        Assert.True(OfficeRasterImageDecoder.TryDecodeFrames(tagged,
            new OfficeRasterDecodeOptions { ApplyExifOrientation = false }, out var storedFrames));
        OfficeRasterImage stored = storedFrames![0].Image;
        Assert.Equal(3, stored.Width);
        Assert.Equal(2, stored.Height);
        Assert.Equal(2, display!.Width);
        Assert.Equal(3, display.Height);
        Assert.Equal(OfficeRasterTransforms.AutoOrient(stored, 6).GetPixels(), display.GetPixels());

        if (format == OfficeImageExportFormat.Jpeg) {
            var direct = OfficeJpegCodec.Decode(tagged,
                new OfficeJpegDecodeOptions(false, false, ignoreExifOrientation: true));
            Assert.Equal(stored.GetPixels(), direct.GetPixels());
        }
    }
}
