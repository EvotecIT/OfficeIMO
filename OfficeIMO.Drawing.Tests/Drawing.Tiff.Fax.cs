using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffFaxTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffFax");
    private static int Sample(int x, int y) => y % 7 == 0 ? 0 : y % 7 == 1 ? 1 : y % 7 == 2 ? x & 1 : (x / 13 + y / 2) & 1;

    [Fact]
    public void IndependentFaxStripsAndTilesRetainRunsPolarityAndFillOrder() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, fields[0]));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, fields[0], out _), fields[0]);
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), fields[0]);
            bool whiteIsZero = fields[3] == "0";
            for (int y = 0; y < 19; y++) for (int x = 0; x < 83; x++) {
                int value = Sample(x, y);
                byte gray = (byte)((whiteIsZero ? 1 - value : value) * 255);
                Assert.True(image!.GetPixel(x, y) == OfficeColor.FromRgb(gray, gray, gray), $"{fields[0]} {x},{y}");
            }
        }
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void TruncatedFaxPayloadIsRejectedByValidationAndDecode(int compression) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"c{compression}-o0-p0-be0-tile0-f1.tif"));
        int entry = Entry(bytes, 279), offset = BitConverter.ToInt32(bytes, entry + 8);
        BitConverter.GetBytes(1).CopyTo(bytes, offset);
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "fax.tif", out _));
    }

    [Theory]
    [InlineData(3, 292)]
    [InlineData(4, 293)]
    public void UncompressedFaxExtensionDoesNotSilentlyDecodeAsNormalFax(int compression, int tag) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"c{compression}-o0-p0-be0-tile0-f1.tif"));
        bytes[Entry(bytes, tag) + 8] = 2;
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "fax.tif", out _));
    }

    [Fact]
    public void UnsignedSingleStripSentinelAlsoPreservesByteSamples() {
        byte[] bytes = OfficeTiffCodec.Encode(new OfficeRasterImage(3, 2, OfficeColor.Blue));
        BitConverter.GetBytes(uint.MaxValue).CopyTo(bytes, Entry(bytes, 278) + 8);
        Assert.True(OfficeImageReader.TryValidateContent(bytes, "single-strip.tif", out _));
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image));
        Assert.Equal(OfficeColor.Blue, image!.GetPixel(2, 1));
    }

    private static int Entry(byte[] bytes, int tag) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) == tag) return entry;
        }
        throw new InvalidOperationException("Missing fixture tag.");
    }
}
