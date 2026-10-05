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

    [Fact]
    public void IndependentExtraSamplesRetainColorAndAlpha() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegAlpha");
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
    [Fact]
    public void IndependentLowAlphaSamplesRetainAlphaAndVisibleCompositing() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "TestAssets", "TiffJpegLowAlpha");
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
                Assert.Equal(expected[p + 3], actual.A);
                byte[] channels = { actual.R, actual.G, actual.B };
                foreach (int background in new[] { 0, 255 }) for (int c = 0; c < 3; c++) {
                    double visible = (channels[c] * actual.A + background * (255 - actual.A)) / 255D;
                    double reference = (expected[p + c] * expected[p + 3] + background * (255 - expected[p + 3])) / 255D;
                    Assert.True(Math.Abs(visible - reference) <= 3, $"{name} {x},{y}: composite {visible} != {reference}");
                }
            }
        }
    }

}
