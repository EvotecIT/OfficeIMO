using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    [InlineData(4)]
    public void ImageReaderReportsJpegComponentCountWithoutClaimingColorSpace(int components) {
        // Header-only evidence: identification must not imply that the pixel payload is valid.
        byte[] header = new byte[12 + components * 3];
        header[0] = 255; header[1] = 216; header[2] = 255; header[3] = 192;
        header[5] = (byte)(8 + components * 3); header[6] = 8;
        header[8] = 3; header[10] = 2; header[11] = (byte)components;
        Assert.True(OfficeImageReader.TryIdentifyByContent(header, null, out var info));
        Assert.Equal(components, info.JpegComponentCount);
        Assert.Null(info.TiffPhotometricInterpretation);
        Assert.False(OfficeRasterImageDecoder.TryDecode(header, out _));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImageReaderReportsTiffPhotometricWithoutInferringMissingOrAmbiguousTags(bool big, bool little) {
        byte[] bytes = big ? CreateBigTiff(little, 11, 9, 300) : CreateClassicTiff(little, 11, 9, 300, 1, 300, 1);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var absent));
        Assert.Null(absent.TiffPhotometricInterpretation);
        int entry = big ? 104 : 58;
        int value = big ? 12 : 8;
        WriteUInt16(bytes, entry, 262, little);
        foreach (ushort photometric in new ushort[] { 0, 1, 2, 3, 5, 6 }) {
            WriteUInt16(bytes, entry + value, photometric, little);
            Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var info));
            Assert.Equal((int)photometric, info.TiffPhotometricInterpretation);
            Assert.Null(info.JpegComponentCount);
        }
        // A LONG is not the specified SHORT field type, even when the number looks like RGB.
        WriteUInt16(bytes, entry + 2, 4, little);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var invalid));
        Assert.Null(invalid.TiffPhotometricInterpretation);
        WriteUInt16(bytes, entry + 2, 3, little);
        int duplicate = big ? 84 : 46;
        if (big) WriteShortEntry(bytes, duplicate, 262, 2, little);
        else WriteClassicShortEntry(bytes, duplicate, 262, 2, little);
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var ambiguous));
        Assert.Null(ambiguous.TiffPhotometricInterpretation);
    }
}
