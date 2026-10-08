using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffLosslessPrecisionTests {
    [Theory]
    [InlineData("huffman")]
    [InlineData("arithmetic")]
    public void NativePrecisionSurvivesTiffWrapping(string coding) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLosslessPrecision");
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            if (!fields[0].StartsWith(coding, StringComparison.Ordinal)) continue;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, fields[0]));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, fields[0], out _), fields[0] + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), fields[0]);
            int width = int.Parse(fields[4]), height = int.Parse(fields[5]);
            Assert.Equal((width, height), (image!.Width, image.Height));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, fields[0] + ".rgba"));
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int at = (y * width + x) * 4;
                var pixel = image.GetPixel(x, y);
                Assert.True(pixel.R == expected[at] && pixel.G == expected[at + 1] &&
                    pixel.B == expected[at + 2] && pixel.A == expected[at + 3], $"{fields[0]} at {x},{y}");
            }
        }
    }
    [Theory]
    [InlineData(2)]
    [InlineData(9)]
    [InlineData(15)]
    public void ExpandedPrecisionHonorsLimitsAndDoesNotAdmitOtherProcesses(int bits) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLosslessPrecision");
        foreach (string coding in new[] { "huffman", "arithmetic" }) {
            string name = coding == "huffman" ? $"huffman-p{bits}-c1-d1-t0-s1.jpg.le.tif" :
                $"arithmetic-b{bits}-c1-p1-t0-r19.jpg.le.tif";
            byte[] original = File.ReadAllBytes(Path.Combine(corpus, name));
            Assert.False(OfficeRasterImageDecoder.TryDecode(original,
                new OfficeRasterDecodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - original.Length }, out _, out _));
            using var cancellation = new System.Threading.CancellationTokenSource();
            cancellation.Cancel();
            Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(original,
                new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));

            byte[] dct = (byte[])original.Clone();
            int strip = ReadTag(original, 273), marker = coding == "huffman" ? 0xC3 : 0xCB;
            int frame = -1;
            for (int i = strip; i + 1 < original.Length; i++)
                if (original[i] == 255 && original[i + 1] == marker) { frame = i + 1; break; }
            Assert.True(frame >= 0);
            dct[frame] = (byte)(coding == "huffman" ? 0xC1 : 0xC9);
            AssertRejected(dct); // DCT remains limited to eight/twelve bits.
            byte[] mismatch = (byte[])original.Clone();
            mismatch[frame + 3] = (byte)(bits + 1);
            AssertRejected(mismatch);
            foreach (int compression in new[] { 1, 6 }) {
                byte[] other = (byte[])original.Clone();
                int at = FindTag(other, 259);
                other[at + 8] = (byte)compression;
                other[at + 9] = 0;
                AssertRejected(other);
            }
        }
    }

    private static void AssertRejected(byte[] bytes) {
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "source.tif", out _));
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
    }

    // The selected malformed-input seeds are little-endian, single-strip files.
    private static int ReadTag(byte[] bytes, int tag) => BitConverter.ToInt32(bytes, FindTag(bytes, tag) + 8);

    private static int FindTag(byte[] bytes, int tag) {
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int at = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, at) == tag) return at;
        }
        throw new InvalidDataException($"Missing TIFF tag {tag}.");
    }
}
