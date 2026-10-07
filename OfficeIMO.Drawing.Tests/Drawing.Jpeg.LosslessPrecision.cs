using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegLosslessPrecisionTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegLosslessPrecision");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EveryLosslessPrecisionMatchesIndependentDecodedSamples(bool highQuality) {
        foreach (string row in File.ReadLines(Path.Combine(Corpus, "manifest.csv")).Skip(1)) {
            string[] f = row.Split(',');
            byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, f[0]));
            byte[] words = File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".raw"));
            int precision = int.Parse(f[1]), channels = int.Parse(f[2]), maximum = (1 << precision) - 1;
            Assert.True(OfficeImageReader.TryValidateContent(jpeg, f[0], out _), f[0]);
            Assert.True(OfficeJpegCodec.TryDecode(jpeg, out var image, new OfficeJpegDecodeOptions(highQualityChroma: highQuality)), f[0]);
            Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out byte[] components,
                out int width, out int height, out int count, options: new OfficeJpegDecodeOptions(highQualityChroma: highQuality)), f[0]);
            Assert.Equal((17, 11, channels), (width, height, count));
            for (int i = 0; i < components.Length; i++) {
                int native = words[i * 2] | words[i * 2 + 1] << 8;
                byte expected = (byte)((native * 255L + maximum / 2) / maximum);
                Assert.True(expected == components[i], $"{f[0]} component {i}: expected {expected}, actual {components[i]}");
            }
            byte[]? colorReference = f[8] == "ycc" ? File.ReadAllBytes(Path.Combine(Corpus, f[0] + ".rgba")) : null;
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                int at = (y * width + x) * channels;
                var pixel = image!.GetPixel(x, y);
                if (colorReference != null) {
                    int p = (y * width + x) * 4;
                    Assert.True(Math.Abs(pixel.R - colorReference[p]) <= 1 && Math.Abs(pixel.G - colorReference[p + 1]) <= 1 &&
                        Math.Abs(pixel.B - colorReference[p + 2]) <= 1 && pixel.A == 255, $"{f[0]} color at {x},{y}");
                } else Assert.Equal((components[at], components[at + (channels == 1 ? 0 : 1)],
                    components[at + (channels == 1 ? 0 : 2)], (byte)255), (pixel.R, pixel.G, pixel.B, pixel.A));
            }
        }
    }

    [Theory]
    [InlineData(2)]
    [InlineData(9)]
    [InlineData(12)]
    [InlineData(15)]
    public void ExpandedPrecisionsHonorCancellationAndRetainedMemory(int precision) {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(Corpus, $"p{precision}-c3-d1-t0-s1.jpg"));
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out _, out _, out _, out _, cancellationToken: cancellation.Token));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false, out _, out _, out _, out _,
            retainedManagedBytes: OfficeRasterGuards.MaximumDecodedBytes - jpeg.Length));
    }
}
