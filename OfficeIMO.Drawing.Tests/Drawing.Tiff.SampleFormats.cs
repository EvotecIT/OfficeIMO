using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingTiffSampleFormatTests {
    [Fact]
    public void UnspecifiedExtraSampleIsIgnoredInsteadOfTreatedAsAlpha() {
        var source = new OfficeRasterImage(1, 1, OfficeColor.FromRgba(128, 64, 32, 64));
        byte[] bytes = OfficeTiffCodec.Encode(source);
        SetShortTag(bytes, 338, 0);
        Assert.True(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var image));
        Assert.Equal(OfficeColor.FromRgb(128, 64, 32), image!.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void UnsupportedSampleEncodingsRejectAcrossValidationAndSelectedDecode(int sampleFormat) {
        // Independent libtiff RGB Deflate fixture with per-channel SampleFormat.
        byte[] bytes = Convert.FromBase64String("SUkqABYAAAB4nPvPwMD4nwEACAECAA8AAAEDAAEAAAACAAAAAQEDAAEAAAABAAAAAgEDAAMAAADQAAAAAwEDAAEAAAAIAAAABgEDAAEAAAACAAAACgEDAAEAAAABAAAAEQEEAAEAAAAIAAAAEgEDAAEAAAABAAAAFQEDAAEAAAADAAAAFgEDAAEAAAABAAAAFwEEAAEAAAAOAAAAHAEDAAEAAAABAAAAKAEDAAEAAAACAAAAPQEDAAEAAAACAAAAUwEDAAMAAADWAAAAAAAAAAgACAAIAAEAAQABAA==");
        SetShortTag(bytes, 339, sampleFormat);
        Assert.False(OfficeTiffCodec.TryDecodePage(bytes, 0, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
    }

    private static void SetShortTag(byte[] bytes, int tag, int value) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int index = 0; index < count; index++) {
            int entry = ifd + 2 + index * 12;
            if (BitConverter.ToUInt16(bytes, entry) != tag) continue;
            int target = BitConverter.ToInt32(bytes, entry + 4) * 2 <= 4 ? entry + 8 : BitConverter.ToInt32(bytes, entry + 8) + 2;
            bytes[target] = (byte)value; bytes[target + 1] = (byte)(value >> 8); return;
        }
        throw new InvalidOperationException("Fixture does not contain the requested TIFF tag.");
    }
}
