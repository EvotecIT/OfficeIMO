using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffExtraSamplesTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffExtraSamples");

    [Fact]
    public void IndependentSamplesSelectDeclaredAlphaAndIgnoreUnspecifiedChannels() {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, f[0]));
            byte[] expected = File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".rgba"));
            int width = int.Parse(f[5]), height = int.Parse(f[6]), tolerance = int.Parse(f[7]);
            Assert.True(OfficeImageReader.TryValidateContent(bytes, f[0], out _), f[0] + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), f[0] + " decode");
            Assert.Equal(width, image!.Width); Assert.Equal(height, image.Height);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int p = (y * width + x) * 4; var actual = image.GetPixel(x, y);
                int delta = Math.Max(Math.Abs(actual.R - expected[p]), Math.Max(Math.Abs(actual.G - expected[p + 1]), Math.Abs(actual.B - expected[p + 2])));
                Assert.True(delta <= tolerance && Math.Abs(actual.A - expected[p + 3]) <= 1,
                    $"{f[0]} {x},{y}: RGB {delta}, alpha {actual.A}/{expected[p + 3]}");
            }
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void AmbiguousOrUnknownExtraSampleMeaningIsRejected(int extraKind) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Corpus, "n-4-l0-e100-p0.tif"));
        int ifd = BitConverter.ToInt32(bytes, 4), count = BitConverter.ToUInt16(bytes, ifd);
        for (int i = 0; i < count; i++) {
            int entry = ifd + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) != 338) continue;
            int offset = BitConverter.ToInt32(bytes, entry + 8);
            BitConverter.GetBytes((ushort)extraKind).CopyTo(bytes, offset + 2);
        }
        Assert.False(OfficeImageReader.TryValidateContent(bytes, "extra.tif", out _));
        Assert.False(OfficeTiffCodec.TryDecode(bytes, out _));
    }
}
