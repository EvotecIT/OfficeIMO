using System;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpDecodeInspectionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StaticLosslessDecodeRetainsPixelsAndIndependentRequestOwnership(bool transparent) {
        OfficeRasterImage source = CreateImage(transparent);
        byte[] expected = source.GetPixels();
        byte[] encoded = OfficeWebpCodec.Encode(source);
        var caller = new UnexpectedCodec();

        Assert.True(OfficeRasterContainerInspector.TryInspect(encoded, out var container));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { ImageCodec = caller }, out var first, out var info));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var second));
        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.NotSame(first, second);
        Assert.Equal(expected, first!.GetPixels());
        Assert.Equal(expected, second!.GetPixels());
        Assert.Equal(OfficeImageFormat.Webp, info.Format);
        Assert.Equal(1, info.FrameCount);
        Assert.Equal(0, info.SelectedFrameIndex);
        Assert.True(info.Succeeded);
        Assert.False(info.IsAnimated);
        Assert.False(info.UsedCallerCodec);
        Assert.Null(info.Diagnostic);
        Assert.NotNull(container);
        var inspected = info.Container;
        Assert.NotNull(inspected);
        Assert.Equal(container.CanvasWidth, inspected.CanvasWidth);
        Assert.Equal(container.CanvasHeight, inspected.CanvasHeight);
        Assert.Equal(container.Count, inspected.Count);

        first.Fill(OfficeColor.White);
        Assert.Equal(expected, second.GetPixels());
        Assert.Equal(expected, source.GetPixels());
    }

    [Fact]
    public void StaticLosslessDecodeRejectsAnUnavailableFrameBeforeReturningPixels() {
        byte[] encoded = OfficeWebpCodec.Encode(CreateImage(false));
        Assert.False(OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { FrameIndex = 1 }, out var image, out var info));
        Assert.Null(image);
        Assert.False(info.Succeeded);
        Assert.Equal(1, info.FrameCount);
        Assert.Equal(1, info.SelectedFrameIndex);
        Assert.NotNull(info.Container);
    }

    [Fact]
    public void MissingLosslessCodesKeepTheInspectionFailureAndCannotReachACallerCodec() {
        byte[] encoded = OfficeWebpCodec.Encode(CreateImage(false));
        int cursor = 12;
        while (Encoding.ASCII.GetString(encoded, cursor, 4) != "VP8L") {
            int length = BitConverter.ToInt32(encoded, cursor + 4);
            cursor += 8 + length + (length & 1);
        }
        byte[] malformed = {
            (byte)'R', (byte)'I', (byte)'F', (byte)'F', 18, 0, 0, 0,
            (byte)'W', (byte)'E', (byte)'B', (byte)'P',
            (byte)'V', (byte)'P', (byte)'8', (byte)'L', 5, 0, 0, 0,
            0, 0, 0, 0, 0, 0
        };
        Array.Copy(encoded, cursor + 8, malformed, 20, 5);
        Assert.True(OfficeImageReader.TryIdentifyByContent(malformed, null, out _));
        Assert.False(OfficeRasterContainerInspector.TryInspect(malformed, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(malformed,
            new OfficeRasterDecodeOptions { ImageCodec = new UnexpectedCodec() }, out var image, out var info));
        Assert.Null(image);
        Assert.False(info.Succeeded);
        Assert.Equal(OfficeImageFormat.Webp, info.Format);
        Assert.Equal(0, info.FrameCount);
        Assert.Null(info.Container);
    }

    [Fact]
    public void StaticLosslessDecodeRetainsResourceAndCallerCancellationLimits() {
        byte[] encoded = OfficeWebpCodec.Encode(CreateImage(true));
        foreach (var options in new[] {
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 13 * 7 - 1 },
            new OfficeRasterDecodeOptions { MaximumEncodedBytes = encoded.Length - 1 },
            new OfficeRasterDecodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes }
        }) {
            Assert.False(OfficeRasterImageDecoder.TryDecode(encoded, options, out var image, out var info));
            Assert.Null(image);
            Assert.False(info.Succeeded);
        }
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var failure = Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
        Assert.Equal(cancellation.Token, failure.CancellationToken);
    }

    private static OfficeRasterImage CreateImage(bool transparent) {
        var image = new OfficeRasterImage(13, 7);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgba((byte)(x * 19), (byte)(y * 31),
                    (byte)(x * 7 + y * 11), transparent ? (byte)((x + y) % 3 * 127) : (byte)255));
            }
        }
        return image;
    }

    private sealed class UnexpectedCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) =>
            throw new InvalidOperationException("Managed lossless WebP must not reach a caller codec.");
    }
}
