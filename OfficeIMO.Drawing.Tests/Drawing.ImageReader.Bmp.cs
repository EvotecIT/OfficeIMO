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
    [InlineData(16, 6, -2)]
    [InlineData(32, 6, -2)]
    public void OfficeImageReaderRejectsBmpCompressionThatDoesNotMatchBitDepth(
        int bitsPerPixel,
        int compression,
        int height) {
        byte[] bmp = CreateBmpInfoHeader(bitsPerPixel, compression, height);

        Assert.False(OfficeImageReader.TryIdentifyByContent(bmp, fileName: null, out _));
    }

    [Fact]
    public void RasterContainerInspectionRejectsBmpOutsideTheManagedDecoderSubset() {
        byte[] bmp = CreateCompleteIndexedBmp8();

        Assert.True(OfficeImageReader.TryIdentifyByContent(bmp, fileName: null, out OfficeImageInfo info));
        Assert.Equal(OfficeImageFormat.Bmp, info.Format);
        Assert.False(OfficeBmpReader.TryDecode(bmp, out _));
        Assert.False(OfficeRasterContainerInspector.TryInspect(bmp, out _));
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

    private static byte[] CreateCompleteIndexedBmp8() {
        const int pixelOffset = 62;
        const int fileSize = 66;
        var bmp = new byte[fileSize];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteInt32LittleEndian(bmp, 2, fileSize);
        WriteInt32LittleEndian(bmp, 10, pixelOffset);
        WriteInt32LittleEndian(bmp, 14, 40);
        WriteInt32LittleEndian(bmp, 18, 1);
        WriteInt32LittleEndian(bmp, 22, 1);
        WriteUInt16LittleEndian(bmp, 26, 1);
        WriteUInt16LittleEndian(bmp, 28, 8);
        WriteInt32LittleEndian(bmp, 34, 4);
        WriteInt32LittleEndian(bmp, 46, 2);
        bmp[54] = 0;
        bmp[55] = 0;
        bmp[56] = 0;
        bmp[58] = 255;
        bmp[59] = 255;
        bmp[60] = 255;
        bmp[pixelOffset] = 1;
        return bmp;
    }
}
