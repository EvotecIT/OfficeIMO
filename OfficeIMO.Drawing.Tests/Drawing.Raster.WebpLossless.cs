using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Lossless VP8L files written by libwebp, compared with libwebp's own decode. Together they reach
/// prefix codes longer than eight bits, the color cache, every palette packing, overlapping
/// backreferences, and the predictor, color and subtract-green transforms.
/// </summary>
public sealed class DrawingWebpLosslessTests {
    [Theory]
    [InlineData("lossless-noise", 61, 47)]
    [InlineData("lossless-smooth", 128, 48)]
    [InlineData("lossless-smooth-fast", 96, 64)]
    [InlineData("lossless-palette-2", 53, 29)]
    [InlineData("lossless-palette-4", 53, 29)]
    [InlineData("lossless-palette-16", 53, 29)]
    [InlineData("lossless-palette-64", 53, 29)]
    [InlineData("lossless-repeat", 97, 31)]
    public void IndependentLosslessFilesDecodeExactly(string name, int width, int height) {
        byte[] bytes = ReadFixture(name + ".webp");
        byte[] expected = ReadFixture(name + ".rgba");

        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image));
        Assert.Equal((width, height), (image!.Width, image.Height));
        Assert.Equal(expected, image.GetPixels());
    }

    [Theory]
    [InlineData("lossless-noise")]
    [InlineData("lossless-palette-4")]
    [InlineData("lossless-repeat")]
    public void TruncatedLosslessPayloadIsRejected(string name) {
        byte[] bytes = ReadFixture(name + ".webp");
        // Drop the last quarter of the VP8L payload and repair both RIFF sizes, so only the
        // entropy-coded data is short.
        int length = bytes.Length - (bytes.Length - 20) / 4;
        byte[] truncated = bytes.AsSpan(0, length).ToArray();
        BitConverter.GetBytes(length - 8).CopyTo(truncated, 4);
        BitConverter.GetBytes(length - 20).CopyTo(truncated, 16);

        Assert.False(OfficeRasterImageDecoder.TryDecode(truncated, out _));
    }

    private static byte[] ReadFixture(string name) {
        using Stream input = typeof(DrawingWebpLosslessTests).Assembly.GetManifestResourceStream(
            "OfficeIMO.Drawing.Tests.TestAssets.WebpLossless." + name)!;
        Assert.NotNull(input);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }
}
