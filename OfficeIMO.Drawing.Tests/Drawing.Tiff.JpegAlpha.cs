using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TiffJpegAlphaTests {
    [Fact]
    public void RawExtraChannelCountsTowardTheJpegWorkingSetLimit() {
        long retained = OfficeRasterGuards.MaximumDecodedBytes - 64L * 1024 - 4L * 1024 * 1024;
        Assert.True(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, outputComponents: 4));
        Assert.False(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, outputComponents: 5));
    }

    [Theory]
    [InlineData("TiffJpegAlpha")]
    [InlineData("TiffJpegArithmeticAlpha")]
    public void IndependentExtraSamplesRetainColorAndAlpha(string folder) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(',');
            string name = fields[0];
            // Unassociation amplifies the independently rounded JPEG/color samples.
            int tolerance = fields[7] == "1" ? 6 : 3;
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, name));
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".rgba"));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), name + " decode");
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                var actual = image!.GetPixel(x, y); int p = (y * 35 + x) * 4;
                int error = Math.Max(Math.Abs(actual.R - expected[p]), Math.Max(Math.Abs(actual.G - expected[p + 1]), Math.Abs(actual.B - expected[p + 2])));
                Assert.True(error <= tolerance, $"{name} {x},{y}: RGB delta {error}");
                Assert.True(Math.Abs(actual.A - expected[p + 3]) <= 1, $"{name} {x},{y}: alpha {actual.A} != {expected[p + 3]}");
            }
        }
    }
    [Theory]
    [InlineData("TiffJpegLowAlpha")]
    [InlineData("TiffJpegArithmeticLowAlpha")]
    [InlineData("TiffJpegArithmetic12")]
    [InlineData("TiffJpegArithmeticLosslessColor")]
    [InlineData("TiffJpegArithmeticLosslessChroma")]
    [InlineData("TiffJpegArithmeticExtra")]
    [InlineData("TiffJpegArithmeticMultiscan")]
    public void IndependentLowAlphaSamplesRetainAlphaAndVisibleCompositing(string folder) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder);
        foreach (string row in File.ReadLines(Path.Combine(corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            byte[] expected = File.ReadAllBytes(Path.Combine(corpus, name + ".rgba"));
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, name));
            Assert.True(OfficeImageReader.TryValidateContent(bytes, name, out _), name + " validation");
            Assert.True(OfficeTiffCodec.TryDecode(bytes, out var image), name + " decode");
            Assert.Equal(35, image!.Width);
            Assert.Equal(19, image.Height);
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
                var actual = image.GetPixel(x, y); int p = (y * image.Width + x) * 4;
                Assert.True(Math.Abs(expected[p + 3] - actual.A) <= (folder == "TiffJpegArithmetic12" ? 1 : 0), $"{name} {x},{y}: alpha {actual.A} != {expected[p + 3]}");
                byte[] channels = { actual.R, actual.G, actual.B };
                foreach (int background in new[] { 0, 255 }) for (int c = 0; c < 3; c++) {
                    double visible = (channels[c] * actual.A + background * (255 - actual.A)) / 255D;
                    double reference = (expected[p + c] * expected[p + 3] + background * (255 - expected[p + 3])) / 255D;
                    Assert.True(Math.Abs(visible - reference) <= 3, $"{name} {x},{y}: composite {visible} != {reference}");
                }
            }
        }
    }

    [Theory]
    [InlineData("TiffJpegArithmeticAlpha", 48)]
    [InlineData("TiffJpegArithmeticLowAlpha", 32)]
    [InlineData("TiffJpegArithmetic12", 64)]
    [InlineData("TiffJpegArithmeticLosslessColor", 24)]
    [InlineData("TiffJpegArithmeticExtra", 6)]
    [InlineData("TiffJpegArithmeticMultiscan", 18)]
    public void ArithmeticCmykAlphaMatchesIndependentProfiledCompositing(string folder, int count) {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", folder);
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(corpus, "..", "IccColorCorpus", "littlecms-cmyk-lut.icc")), out var profile));
        string[] references = Directory.GetFiles(corpus, "*.icc-rgba");
        Assert.Equal(count, references.Length);
        foreach (string reference in references) {
            byte[] expected = File.ReadAllBytes(reference);
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(File.ReadAllBytes(reference.Substring(0, reference.Length - ".icc-rgba".Length)),
                profile!, new(), out var image), reference);
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                var actual = image!.GetPixel(x, y); int at = (y * 35 + x) * 4;
                Assert.True(Math.Abs(expected[at + 3] - actual.A) <= (folder == "TiffJpegArithmetic12" ? 1 : 0), $"{reference} {x},{y}: alpha {actual.A} != {expected[at + 3]}");
                byte[] components = { actual.R, actual.G, actual.B };
                foreach (int background in new[] { 0, 255 }) for (int c = 0; c < 3; c++) {
                    double rendered = (components[c] * actual.A + background * (255 - actual.A)) / 255D;
                    double native = (expected[at + c] * expected[at + 3] + background * (255 - expected[at + 3])) / 255D;
                    Assert.True(Math.Abs(rendered - native) <= 3, $"{reference} {x},{y}: {rendered} != {native}");
                }
            }
        }
    }

}
