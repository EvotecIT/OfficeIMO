using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingJpegHuffmanRefinementTests {
    private static string Corpus => Path.Combine(AppContext.BaseDirectory, "TestAssets", "JpegHuffmanRefinement");

    // libjpeg-turbo encoded each crop as baseline and as successive approximation with
    // the same quantization, so refinement must rebuild the baseline coefficients.
    [Theory]
    [InlineData("color-420")]
    [InlineData("color-444-restart")]
    [InlineData("gray")]
    public void SuccessiveApproximationDecodesToTheBaselinePixels(string name) {
        byte[] baseline = File.ReadAllBytes(Path.Combine(Corpus, name + "-baseline.jpg"));
        byte[] progressive = File.ReadAllBytes(Path.Combine(Corpus, name + "-progressive.jpg"));
        foreach (bool highQualityChroma in new[] { false, true }) {
            var options = new OfficeJpegDecodeOptions(highQualityChroma);
            OfficeRasterImage expected = OfficeJpegCodec.Decode(baseline, options);
            OfficeRasterImage actual = OfficeJpegCodec.Decode(progressive, options);
            Assert.Equal((120, 72), (actual.Width, actual.Height));
            Assert.True(expected.GetPixels().SequenceEqual(actual.GetPixels()), $"{name} highQualityChroma={highQualityChroma}");
        }
    }
}
