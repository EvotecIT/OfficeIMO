using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpLosslessCancellationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ByteArrayOverloadsPreserveCompressionAndEveryRgbaSample(bool density, bool transparency) {
        OfficeRasterImage image = CreateRepeatedImage(transparency);
        byte[] original = image.GetPixels();
        var options = new OfficeWebpEncodeOptions {
            WritePhysicalResolution = density, DpiX = 144, DpiY = 120
        };
        byte[] expected = density ? OfficeWebpCodec.Encode(image, 144, 120) : OfficeWebpCodec.Encode(image);
        using var cancellation = new CancellationTokenSource();

        Assert.Equal(expected, OfficeWebpCodec.Encode(image, options));
        Assert.Equal(expected, OfficeWebpCodec.Encode(image, options, CancellationToken.None));
        Assert.Equal(expected, OfficeWebpCodec.Encode(image, options, cancellation.Token));
        Assert.True(expected.Length < original.Length / 4);
        Assert.True(OfficeWebpCodec.TryDecode(expected, out OfficeRasterImage? decoded));
        Assert.Equal(original, decoded!.GetPixels());
        Assert.Equal(original, image.GetPixels());
        if (density) {
            OfficeImageInfo info = OfficeImageReader.Identify(expected);
            Assert.Equal(144D, info.DpiX, 6);
            Assert.Equal(120D, info.DpiY, 6);
        }
    }

    [Fact]
    public void CancelledByteArrayEncodingPreservesTheSource() {
        OfficeRasterImage image = CreateRepeatedImage(transparency: true);
        byte[] original = image.GetPixels();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeWebpCodec.Encode(image, new OfficeWebpEncodeOptions(), cancellation.Token));
        Assert.Equal(original, image.GetPixels());
        Assert.True(OfficeWebpCodec.TryDecode(OfficeWebpCodec.Encode(image), out OfficeRasterImage? decoded));
        Assert.Equal(original, decoded!.GetPixels());
    }

    [Fact]
    public void ByteArrayWorkingSetIncludesCallerBuffersBeforeAllocatingCompressionScratch() {
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        var options = new OfficeWebpEncodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes };

        Assert.Throws<ArgumentException>(() => OfficeWebpCodec.Encode(image, options));
        Assert.Throws<ArgumentException>(() => OfficeWebpCodec.Encode(image, options, CancellationToken.None));
    }

    [Fact]
    public void OptionalCompressionFallsBackToLiteralWhenOnlyTheLiteralWorkingSetFits() {
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        var options = new OfficeWebpEncodeOptions {
            RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - 64L * 1024L
        };
        using var stream = new MemoryStream();
        OfficeWebpCodec.EncodeTo(image, stream);
        byte[] expected = stream.ToArray();

        Assert.Equal(expected, OfficeWebpCodec.Encode(image, options));
        Assert.Equal(expected, OfficeWebpCodec.Encode(image, options, CancellationToken.None));
        Assert.True(OfficeWebpCodec.TryDecode(expected, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.White, decoded!.GetPixel(0, 0));
    }

    private static OfficeRasterImage CreateRepeatedImage(bool transparency) {
        var image = new OfficeRasterImage(257, 193);
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgba((byte)(x / 8 * 17), (byte)(y / 8 * 11),
                    (byte)(x / 8 * 17 + y / 8 * 11),
                    transparency ? (byte)((x + y) % 5 == 0 ? 0 : 160) : (byte)255));
            }
        }
        return image;
    }
}
