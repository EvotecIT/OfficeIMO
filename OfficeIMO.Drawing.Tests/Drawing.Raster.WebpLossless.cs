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

        // The caller-owned buffer is the one the packed pixels were decoded into.
        Assert.True(OfficeRasterImageDecoder.TryDecodeRgba32(bytes, out byte[]? pixels, out int rgbaWidth, out int rgbaHeight));
        Assert.Equal((width, height), (rgbaWidth, rgbaHeight));
        Assert.Equal(expected, pixels);
    }

    [Theory]
    [InlineData("lossless-noise")]
    [InlineData("lossless-palette-4")]
    [InlineData("lossless-repeat")]
    public void TruncatedLosslessPayloadIsRejected(string name) {
        byte[] bytes = ReadFixture(name + ".webp");
        int payload = BitConverter.ToInt32(bytes, 16);

        // The rebuilt container decodes at full length, so the shortened one fails only
        // because its entropy-coded data ends early.
        Assert.True(OfficeRasterImageDecoder.TryDecode(WithVp8lPayloadLength(bytes, payload), out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(WithVp8lPayloadLength(bytes, payload - payload / 4), out _));
    }

    /// <summary>Keeps the first <paramref name="length"/> payload bytes with consistent sizes and the RIFF pad byte.</summary>
    private static byte[] WithVp8lPayloadLength(byte[] bytes, int length) {
        var result = new byte[20 + length + (length & 1)];
        Array.Copy(bytes, result, 20 + length);
        BitConverter.GetBytes(result.Length - 8).CopyTo(result, 4);
        BitConverter.GetBytes(length).CopyTo(result, 16);
        return result;
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
