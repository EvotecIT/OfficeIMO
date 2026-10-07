using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpHuffmanTests {
    // Independently assembled VP8L fixtures have a complete green tree with
    // code lengths 1..9,9 and singleton red/blue/alpha/distance trees. Native
    // decoding also qualifies their expected pixels. The one-pixel stream
    // ends with exactly one available bit; larger streams mix lookup and
    // nine-bit canonical symbols across byte boundaries.
    [Theory]
    [InlineData(1, "UklGRiYAAABXRUJQVlA4TBkAAAAvAAAAAPAASZIgSZIkC0FiUfPIbEM/+l8AAA==")]
    [InlineData(2, "UklGRigAAABXRUJQVlA4TBsAAAAvAQAAAPAASZIgSZIkC0FiUfPIbEM/+l8A/wEA")]
    [InlineData(3, "UklGRigAAABXRUJQVlA4TBwAAAAvAgAAAPAASZIgSZIkC0FiUfPIbEM/+l8A//8B")]
    [InlineData(9, "UklGRiwAAABXRUJQVlA4TB8AAAAvCAAAAPAASZIgSZIkC0FiUfPIbEM/+l8A///1t989AA==")]
    [InlineData(17, "UklGRjIAAABXRUJQVlA4TCUAAAAvEAAAAPAASZIgSZIkC0FiUfPIbEM/+l8A///1t9+9z/9//e0HAA==")]
    public void MixedLookupAndLongCodesPreservePixelsAndShortFinalSymbols(int width, string fixture) {
        byte[] encoded = Convert.FromBase64String(fixture);
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out var decoded));
        Assert.NotNull(decoded);
        Assert.Equal(width, decoded!.Width);
        Assert.Equal(1, decoded.Height);
        byte[] greens = { 0, 9, 8, 1, 7, 2, 6, 3, 4, 5 };
        var expected = new byte[width * 4];
        for (int pixel = 0; pixel < width; pixel++) {
            expected[pixel * 4] = 13;
            expected[pixel * 4 + 1] = greens[pixel % greens.Length];
            expected[pixel * 4 + 2] = 31;
            expected[pixel * 4 + 3] = 255;
        }
        Assert.Equal(expected, decoded.GetPixels());
    }

    [Theory]
    [InlineData("UklGRiYAAABXRUJQVlA4TBkAAAAvAAAAAPAASZIgSZIkC0FiUfPIbEM/+l8AAA==")]
    [InlineData("UklGRigAAABXRUJQVlA4TBsAAAAvAQAAAPAASZIgSZIkC0FiUfPIbEM/+l8A/wEA")]
    [InlineData("UklGRigAAABXRUJQVlA4TBwAAAAvAgAAAPAASZIgSZIkC0FiUfPIbEM/+l8A//8B")]
    [InlineData("UklGRiwAAABXRUJQVlA4TB8AAAAvCAAAAPAASZIgSZIkC0FiUfPIbEM/+l8A///1t989AA==")]
    [InlineData("UklGRjIAAABXRUJQVlA4TCUAAAAvEAAAAPAASZIgSZIkC0FiUfPIbEM/+l8A///1t9+9z/9//e0HAA==")]
    public void TruncatedSymbolDataIsRejectedInsideAValidRiffContainer(string fixture) {
        byte[] encoded = Convert.FromBase64String(fixture);
        int payloadLength = ReadLittleEndian(encoded, 16);
        var truncatedPayload = new byte[payloadLength - 1];
        Buffer.BlockCopy(encoded, 20, truncatedPayload, 0, truncatedPayload.Length);
        Assert.False(OfficeWebpCodec.TryDecode(WrapPayload(truncatedPayload), out var decoded));
        Assert.Null(decoded);
    }

    [Fact]
    public void NonzeroTrailingBitsAreRejectedAfterTheLastPixel() {
        byte[] encoded = Convert.FromBase64String("UklGRiYAAABXRUJQVlA4TBkAAAAvAAAAAPAASZIgSZIkC0FiUfPIbEM/+l8AAA==");
        int payloadLength = ReadLittleEndian(encoded, 16);
        var payload = new byte[payloadLength + 1];
        Buffer.BlockCopy(encoded, 20, payload, 0, payloadLength);
        Assert.True(OfficeWebpCodec.TryDecode(WrapPayload(payload), out _));
        payload[payloadLength] = 1;
        Assert.False(OfficeWebpCodec.TryDecode(WrapPayload(payload), out var decoded));
        Assert.Null(decoded);
    }

    private static byte[] WrapPayload(byte[] payload) {
        var encoded = new byte[20 + payload.Length + (payload.Length & 1)];
        System.Text.Encoding.ASCII.GetBytes("RIFF").CopyTo(encoded, 0);
        WriteLittleEndian(encoded, 4, encoded.Length - 8);
        System.Text.Encoding.ASCII.GetBytes("WEBPVP8L").CopyTo(encoded, 8);
        WriteLittleEndian(encoded, 16, payload.Length);
        payload.CopyTo(encoded, 20);
        return encoded;
    }

    private static int ReadLittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24;

    private static void WriteLittleEndian(byte[] bytes, int offset, int value) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (index * 8));
    }
}
