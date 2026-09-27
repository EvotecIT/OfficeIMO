using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterOptionalCodecTests {
    private const string IndependentVp8 = "UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=";

    [Fact]
    public void OptionalCodecNormalizesInspectedPayloadWithoutMutatingInput() {
        byte[] bytes = Convert.FromBase64String(IndependentVp8);
        byte[] original = (byte[])bytes.Clone();
        var codec = new Codec { MutateInput = true };
        Assert.True(OfficeImagePngConverter.TryConvertToPng(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out byte[] png, out var info));
        Assert.Equal(1, codec.Calls);
        Assert.Equal(original, bytes);
        Assert.True(info.Succeeded);
        Assert.Equal(OfficeImageFormat.Webp, info.Format);
        Assert.True(OfficePngReader.TryDecode(png, out var image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(8, 8));
    }

    [Theory]
    [InlineData(255, 128)]
    [InlineData(256, 1)]
    public void ResourceLimitsRefuseBeforeInvokingOptionalCodec(long pixels, int encodedBytes) {
        var codec = new Codec();
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentVp8),
            new OfficeRasterDecodeOptions { ImageCodec = codec, MaximumDecodedPixels = pixels, MaximumEncodedBytes = encodedBytes },
            out _, out _));
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void OptionalCodecCancellationIsObservedAfterCallback() {
        using var cts = new CancellationTokenSource();
        var codec = new Codec { AfterDecode = cts.Cancel };
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(
            Convert.FromBase64String(IndependentVp8),
            new OfficeRasterDecodeOptions { ImageCodec = codec, CancellationToken = cts.Token }, out _, out _));
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void OptionalCodecCannotReturnDimensionsDifferentFromInspectedContainer() {
        var codec = new Codec { Size = 1 };
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentVp8),
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void MalformedContainerIsRejectedBeforeOptionalCodec() {
        var codec = new Codec();
        byte[] bytes = Convert.FromBase64String(IndependentVp8);
        Array.Resize(ref bytes, bytes.Length - 1);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out _, out _));
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void OptionalCodecFormatFailureReturnsFailedDecode() {
        var codec = new Codec { AfterDecode = () => throw new FormatException("Unsupported coding feature") };
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentVp8),
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void BuiltInDecodeDoesNotInvokeOptionalCodec() {
        var codec = new Codec();
        byte[] bytes = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Equal(OfficeColor.Blue, image!.GetPixel(0, 0));
        Assert.Equal(0, codec.Calls);
    }

    private sealed class Codec : IOfficeRasterImageCodec {
        internal int Calls;
        internal bool MutateInput;
        internal int Size = 16;
        internal Action? AfterDecode;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            Assert.Equal("image/webp", contentType);
            if (MutateInput) Array.Clear(bytes, 0, bytes.Length);
            image = new OfficeRasterImage(Size, Size, OfficeColor.Red);
            AfterDecode?.Invoke();
            return true;
        }
    }
}
