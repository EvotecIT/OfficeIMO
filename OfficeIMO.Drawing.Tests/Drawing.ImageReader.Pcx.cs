using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImageReaderRejectsPcxWithoutEncodedImageData() {
        byte[] pcx = CreateCompletePcx(2, 2);
        Array.Resize(ref pcx, 128);

        Assert.False(OfficeImageReader.TryIdentifyByContent(pcx, "header-only.pcx", out _));
    }

    [Fact]
    public void OfficeImageReaderRejectsPcxWithAnIncompleteRleScanline() {
        byte[] pcx = CreateCompletePcx(2, 2);
        pcx[128] = 0xC4;

        Assert.False(OfficeImageReader.TryIdentifyByContent(pcx, "overrun.pcx", out _));
    }

    private static byte[] CreateCompletePcx(int width, int height) {
        int bytesPerLine = (width + 1) & ~1;
        int imageDataLength = bytesPerLine * height;
        var pcx = new byte[128 + imageDataLength + 769];
        pcx[0] = 0x0A;
        pcx[1] = 0x05;
        pcx[2] = 0x01;
        pcx[3] = 0x08;
        WriteUInt16LittleEndian(pcx, 8, (ushort)(width - 1));
        WriteUInt16LittleEndian(pcx, 10, (ushort)(height - 1));
        WriteUInt16LittleEndian(pcx, 12, 96);
        WriteUInt16LittleEndian(pcx, 14, 96);
        pcx[65] = 1;
        WriteUInt16LittleEndian(pcx, 66, (ushort)bytesPerLine);

        for (int offset = 128; offset < 128 + imageDataLength; offset++) {
            pcx[offset] = 1;
        }

        pcx[128 + imageDataLength] = 0x0C;
        return pcx;
    }
}
