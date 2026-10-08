using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegArithmeticColorTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegArithmeticColor");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdobeCmykAndYcckMatchNativeColorantPolarity(bool fancy) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] fields = row.Split(','); int maximum = (1 << int.Parse(fields[1])) - 1;
            byte[] jpeg = Read(fields[0]);
            byte[] native = Read(fields[0] + (fancy ? ".fancy.cmyk16" : ".nearest.cmyk16"));
            var options = new OfficeJpegDecodeOptions(highQualityChroma: fancy);
            Assert.True(OfficeImageReader.TryValidateContent(jpeg, fields[0], out _), fields[0]);
            Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, null, false, out var samples,
                out int width, out int height, out int channels, options), fields[0]);
            Assert.Equal((35, 19, 4), (width, height, channels));
            Assert.Equal(native.Length / 2, samples.Length);
            int Expected(int at) => 255 - (((native[at * 2] | native[at * 2 + 1] << 8) * 255 + maximum / 2) / maximum);
            for (int i = 0; i < samples.Length; i++)
                Assert.True(Math.Abs(samples[i] - Expected(i)) <= 2, $"{fields[0]} fancy={fancy} sample {i}");
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var raster, options));
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int at = (y * width + x) * 4; var pixel = raster!.GetPixel(x, y);
                int Rgb(int c) => 255 - Math.Min(255, Expected(at + c) + Expected(at + 3));
                Assert.True(Math.Abs(pixel.R - Rgb(0)) <= 4 && Math.Abs(pixel.G - Rgb(1)) <= 4 &&
                    Math.Abs(pixel.B - Rgb(2)) <= 4 && pixel.A == 255, fields[0]);
            }
        }
    }

    [Fact]
    public void ExplicitCmykProfileMatchesNativeColorTransform() {
        Assert.True(OfficeIccColorProfile.TryCreate(File.ReadAllBytes(Path.Combine(Corpus, "..", "IccColorCorpus", "littlecms-cmyk-lut.icc")), out var profile));
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string name = row.Split(',')[0];
            Assert.True(OfficeIccRasterConverter.TryDecodeToSrgb(Read(name), profile!, new(), out var raster), name);
            byte[] expected = Read(name + ".srgb");
            for (int y = 0; y < 19; y++) for (int x = 0; x < 35; x++) {
                int at = (y * 35 + x) * 3; var pixel = raster!.GetPixel(x, y);
                Assert.True(Math.Abs(pixel.R - expected[at]) <= 3 && Math.Abs(pixel.G - expected[at + 1]) <= 3 &&
                    Math.Abs(pixel.B - expected[at + 2]) <= 3 && pixel.A == 255, $"{name} at {x},{y}");
            }
        }
    }

    private static byte[] Read(string name) => File.ReadAllBytes(Path.Combine(Corpus, name));
}
