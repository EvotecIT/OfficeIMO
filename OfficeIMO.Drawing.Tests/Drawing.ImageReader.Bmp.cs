using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(1, 0)]
    [InlineData(4, 0)]
    [InlineData(4, 2)]
    [InlineData(8, 0)]
    [InlineData(8, 1)]
    [InlineData(16, 0)]
    [InlineData(16, 3)]
    [InlineData(16, 6)]
    [InlineData(24, 0)]
    [InlineData(32, 0)]
    [InlineData(32, 3)]
    [InlineData(32, 6)]
    [InlineData(0, 4)]
    [InlineData(0, 5)]
    public void OfficeImageReaderAcceptsSupportedBmpCompressionForMetadata(int bitsPerPixel, int compression) {
        byte[] bmp = CreateBmpInfoHeader(bitsPerPixel, compression, height: 2);

        Assert.True(OfficeImageReader.TryIdentify(bmp, fileName: null, out OfficeImageInfo image));
        Assert.Equal(OfficeImageFormat.Bmp, image.Format);
    }

    [Theory]
    [InlineData(1, 1, 2)]
    [InlineData(4, 1, 2)]
    [InlineData(8, 2, 2)]
    [InlineData(16, 1, 2)]
    [InlineData(24, 1, 2)]
    [InlineData(24, 4, 2)]
    [InlineData(24, 5, 2)]
    [InlineData(24, 99, 2)]
    [InlineData(32, 2, 2)]
    [InlineData(0, 0, 2)]
    [InlineData(4, 2, -2)]
    [InlineData(8, 1, -2)]
    public void OfficeImageReaderRejectsBmpCompressionThatDoesNotMatchBitDepth(
        int bitsPerPixel,
        int compression,
        int height) {
        byte[] bmp = CreateBmpInfoHeader(bitsPerPixel, compression, height);

        Assert.False(OfficeImageReader.TryIdentifyByContent(bmp, fileName: null, out _));
    }

    private static byte[] CreateBmpInfoHeader(int bitsPerPixel, int compression, int height) {
        var bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteInt32LittleEndian(bmp, 14, 40);
        WriteInt32LittleEndian(bmp, 18, 2);
        WriteInt32LittleEndian(bmp, 22, height);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, (ushort)bitsPerPixel);
        WriteInt32LittleEndian(bmp, 30, compression);
        return bmp;
    }
}
