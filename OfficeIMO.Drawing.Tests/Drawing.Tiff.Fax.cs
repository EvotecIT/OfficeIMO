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

    [Fact]
    public void UncompressedExtensionsResumeOrdinaryRunsAndPreserveRows() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffFaxUncompressed");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, name));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".rgb"));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), name + " decode");
            for (int y = 0; y < 12; y++) for (int x = 0; x < 32; x++) {
                byte gray = expected[(y * 32 + x) * 3];
                Assert.True(image!.GetPixel(x, y) == OfficeColor.FromRgb(gray, gray, gray), $"{name} {x},{y}");
            }
        }
    }

    [Fact]
    public void LiteralRowsRemainReferencesForVerticalAndPassModes() {
        string literal = "0000001111";
        string exit = "00000010";
        string first = literal + "00110011" + exit;
        string second = literal + "0011" + exit + "11";
        string third = literal + "000000010" + "000111"; // One white pixel, exit, pass, V(0), V(0).
        byte[] result = OfficeFaxDecoder.Decode(FaxBytes(first + second + third), 8, 3, -1,
            false, false, true, false, 3, default);
        Assert.Equal(new byte[] { 0x33, 0x33, 0x03 }, result);
    }

    [Theory]
    [InlineData("111", 0x33)] // V(0), V(0), V(0).
    [InlineData("01111", 0x3B)] // V(+1), V(0), V(0).
    [InlineData("00011", 0x3F)] // Pass, V(0).
    public void LiteralExitColorCanResumeAtAReferenceChangeOnItsBoundary(string modes, byte expected) {
        string first = "0000001111" + "00110011" + "00000010";
        string second = "0000001111" + "0011" + "00000011" + modes;
        byte[] result = OfficeFaxDecoder.Decode(FaxBytes(first + second + "000000000001000000000001"), 8, 2, -1,
            false, false, true, true, 2, default);
        Assert.Equal(new byte[] { 0x33, expected }, result);
    }

    [Theory]
    [InlineData("0000001000", 8)] // Reserved extension selector.
    [InlineData("0000001111000000000001", 8)] // Invalid literal code.
    [InlineData("00000011111100000010", 1)] // Two black pixels exceed one column.
    [InlineData("0000001111", 8)] // Missing literal data and exit.
    public void MalformedUncompressedFaxFailsWithinItsInputAndRow(string bits, int columns) {
        Assert.Throws<InvalidDataException>(() => OfficeFaxDecoder.Decode(FaxBytes(bits), columns, 1, -1,
            false, false, true, false, 1, default));
    }

    private static byte[] FaxBytes(string bits) {
        bits = bits.PadRight((bits.Length + 7) / 8 * 8, '0');
        return Enumerable.Range(0, bits.Length / 8).Select(i => Convert.ToByte(bits.Substring(i * 8, 8), 2)).ToArray();
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
    public void UncompressedOptionPermitsSegmentsThatUseOnlyNormalFax(int compression, int tag) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, $"c{compression}-o0-p0-be0-tile0-f1.tif"));
        bytes[Entry(bytes, tag) + 8] = 2;
        Assert.True(OfficeTiffCodec.TryDecode(bytes, out _));
        Assert.True(OfficeImageReader.TryValidateContent(bytes, "fax.tif", out _));
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
