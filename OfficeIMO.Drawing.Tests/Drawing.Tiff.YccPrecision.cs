using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffYccPrecisionTests {
    [Theory]
    [InlineData("huffman", false)]
    [InlineData("arithmetic", false)]
    [InlineData("huffman", true)]
    [InlineData("arithmetic", true)]
    public void YccConversionRetainsFractionalColorUntilFinalProjection(string coding, bool profiled) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegYccPrecision");
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", "icc-dci-p3-matrix.icc")), out var profile));
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            if (!f[0].StartsWith(coding, StringComparison.Ordinal)) continue;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, f[0]));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, f[0], out _), f[0]);
            OfficeRasterImage? image;
            Assert.True(profiled ? OfficeIccRasterConverter.TryDecodeToSrgb(bytes, profile!, new(), out image) :
                OfficeTiffCodec.TryDecode(bytes, out image), f[0]);
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, f[0] + (profiled ? ".icc-rgba" : ".rgba")));
            int tolerance = profiled ? 2 : 1;
            for (int y = 0; y < image!.Height; y++) for (int x = 0; x < image.Width; x++) {
                int at = (y * image.Width + x) * 4;
                var p = image.GetPixel(x, y);
                Assert.True(Math.Abs(p.R - expected[at]) <= tolerance && Math.Abs(p.G - expected[at + 1]) <= tolerance &&
                    Math.Abs(p.B - expected[at + 2]) <= tolerance && p.A == 255,
                    $"{f[0]} at {x},{y}: {p.R},{p.G},{p.B} != {expected[at]},{expected[at + 1]},{expected[at + 2]}");
            }
        }
    }
}
