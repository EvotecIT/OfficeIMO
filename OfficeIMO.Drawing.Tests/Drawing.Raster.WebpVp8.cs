using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpVp8Tests {
    // Independently encoded and decoded with Pillow/libwebp (quality 80, method 6).
    // These fixtures exercise directional prediction, coefficients and odd canvas padding.
    [Theory]
    [InlineData("independent-vp8-pattern", 48, 32)]
    [InlineData("independent-vp8-odd-color", 65, 49)]
    public void OwnedDecoderMatchesIndependentPixels(string name, int width, int height) {
        byte[] encoded = ReadFixture(name + ".webp");
        byte[] expected = ReadFixture(name + ".rgba");
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var image));
        Assert.True(OfficeRasterContainerInspector.TryInspect(encoded, out var container));
        Assert.Equal((width, height), (container!.CanvasWidth, container.CanvasHeight));
        Assert.Equal(width, image!.Width);
        Assert.Equal(height, image.Height);
        byte[] actual = image.GetPixels();
        Assert.Equal(expected.Length, actual.Length);
        long totalError = 0;
        for (int index = 0; index < actual.Length; index++) {
            int error = Math.Abs(actual[index] - expected[index]);
            Assert.InRange(error, 0, 3);
            totalError += error;
        }
        Assert.True(totalError <= actual.Length, "Mean channel error exceeds one.");
        Assert.True(OfficeImageReader.TryValidateContent(encoded, null, out _));
    }

    [Theory]
    [InlineData("UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=", 255, 1, 0)]
    [InlineData("UklGRiQAAABXRUJQVlA4IBgAAABQAQCdASoQABAAAUAmJaQABHQAAP4AAAA=", 127, 127, 127)]
    public void OwnedDecoderPreservesIndependentSolidColors(string encoded, int red, int green, int blue) {
        Assert.True(OfficeWebpCodec.TryDecode(Convert.FromBase64String(encoded), out var image));
        Assert.Equal(16, image!.Width);
        Assert.Equal(16, image.Height);
        byte[] pixels = image.GetPixels();
        for (int offset = 0; offset < pixels.Length; offset += 4) {
            Assert.InRange((int)pixels[offset], Math.Max(0, red - 1), Math.Min(255, red + 1));
            Assert.InRange((int)pixels[offset + 1], Math.Max(0, green - 1), Math.Min(255, green + 1));
            Assert.InRange((int)pixels[offset + 2], Math.Max(0, blue - 1), Math.Min(255, blue + 1));
            Assert.Equal(255, pixels[offset + 3]);
        }
    }

    [Fact]
    public void PixelLimitAndCancellationStopOwnedDecode() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        Assert.False(OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 100 }, out _, out _));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
    }

    [Fact]
    public void AggregateDecodeBudgetRejectsLargeCanvasBeforeAllocatingPlanes() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        // 49 million pixels fit the pixel ceiling but planes plus RGBA exceed the shared byte budget.
        encoded[26] = encoded[28] = (byte)(7000 & 255);
        encoded[27] = encoded[29] = (byte)(7000 >> 8);
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
    }

    [Fact]
    public void TruncatedControlPartitionIsRejected() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        Array.Resize(ref encoded, 24);
        encoded[4] = 16;
        encoded[5] = encoded[6] = encoded[7] = 0;
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExhaustedArithmeticPartitionCannotProduceAnImage(bool tokenPartition) {
        byte[] encoded;
        if (tokenPartition) {
            encoded = ReadFixture("independent-vp8-pattern.webp");
            int controlLength = (encoded[20] | encoded[21] << 8 | encoded[22] << 16) >> 5;
            int payloadLength = 10 + controlLength + 2;
            Array.Resize(ref encoded, 20 + payloadLength + (payloadLength & 1));
            Array.Clear(encoded, 30 + controlLength, encoded.Length - 30 - controlLength);
            WriteLength(encoded, 16, payloadLength);
            WriteLength(encoded, 4, encoded.Length - 8);
        } else {
            // Valid container/header, but only two zero bytes in each partition.
            encoded = new byte[] {82,73,70,70,26,0,0,0,87,69,66,80,86,80,56,32,
                14,0,0,0,80,0,0,157,1,42,16,0,16,0,0,0,0,0};
        }
        Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out _));
        Assert.False(OfficeRasterContainerInspector.TryInspect(encoded, out _));
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(encoded, out _));
        Assert.False(OfficeImageReader.TryValidateContent(encoded, null, out _));
    }

    [Fact]
    public void InverseVp8TransformsUseReferenceSigned16IntermediateStorage() {
        // RFC 6386 section 14 stores dequantized coefficients and each transform
        // stage in signed 16-bit arrays. These extreme reference vectors diverge
        // if a stage is retained as an unbounded 32-bit integer.
        var coefficients = new int[16];
        Array.Fill(coefficients, 32767);
        Assert.Equal(new[] {
            -2402, 478, -478, -95, -12062, 2399, -2399, -477,
            12062, -2399, 2399, 477, 2400, -477, 477, 95
        }, OfficeVp8Transform.InverseTransform4x4(coefficients));
        Assert.Equal(new[] {
            -2, 0, 0, 0, 0, 0, 0, 0,
            0, 0, 0, 0, 0, 0, 0, 0
        }, OfficeVp8Transform.InverseWalshTransform4x4(coefficients));
    }

    private static void WriteLength(byte[] bytes, int offset, int value) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (8 * index));
    }

    private static byte[] ReadFixture(string name) {
        using Stream input = typeof(DrawingWebpVp8Tests).Assembly.GetManifestResourceStream(
            "OfficeIMO.Drawing.Tests.TestAssets.WebpVp8." + name)!;
        Assert.NotNull(input);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }
}
