using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterCodecRobustnessTests {
    [Fact]
    public void DeterministicCompressedPayloadMutationsFailClosedWithoutEscapingExceptions() {
        OfficeRasterImage source = CreatePattern(64, 48);
        byte[] tiff = OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.Lzw });
        byte[] webp = OfficeWebpCodec.Encode(source);

        int rejectedTiff = MutateAndDecode(tiff, 32);
        int rejectedWebp = MutateAndDecode(webp, 20);

        Assert.True(rejectedTiff > 0);
        Assert.True(rejectedWebp > 0);
    }

    [Fact]
    public void TruncatedCompressedPayloadsRemainBoundedAndFailClosed() {
        OfficeRasterImage source = CreatePattern(96, 64);
        byte[][] encoded = {
            OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.Lzw }),
            OfficeWebpCodec.Encode(source)
        };

        foreach (byte[] bytes in encoded) {
            for (int divisor = 2; divisor <= 8; divisor++) {
                int length = Math.Max(1, bytes.Length - bytes.Length / divisor);
                var truncated = new byte[length];
                Buffer.BlockCopy(bytes, 0, truncated, 0, length);
                Assert.False(OfficeRasterImageDecoder.TryDecode(truncated, out _));
            }
        }
    }

    [Fact]
    public void BoundedPngAndTiffDecodeObserveCancellationInsideCodecWork() {
        var source = new OfficeRasterImage(4096, 1025, OfficeColor.FromRgba(24, 80, 160, 224));
        byte[][] encoded = {
            OfficePngWriter.Encode(source),
            OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.Lzw,
                Predictor = OfficeTiffPredictor.Horizontal
            })
        };

        foreach (byte[] bytes in encoded) {
            using var cancellation = new CancellationTokenSource();
            cancellation.CancelAfter(TimeSpan.FromMilliseconds(1));
            var options = new OfficeRasterDecodeOptions {
                CancellationToken = cancellation.Token
            };

            Assert.Throws<OperationCanceledException>(() =>
                OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
        }
    }

    private static int MutateAndDecode(byte[] encoded, int firstOffset) {
        int rejected = 0;
        int step = Math.Max(1, (encoded.Length - firstOffset) / 48);
        for (int offset = firstOffset; offset < encoded.Length; offset += step) {
            byte[] mutated = (byte[])encoded.Clone();
            mutated[offset] ^= (byte)(1 << (offset & 7));
            if (!OfficeRasterImageDecoder.TryDecode(mutated, out _)) rejected++;
        }
        return rejected;
    }

    private static OfficeRasterImage CreatePattern(int width, int height) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                image.SetPixel(x, y, OfficeColor.FromRgba(
                    (byte)(x * 13 + y * 3),
                    (byte)(x * 5 + y * 11),
                    (byte)((x ^ y) * 7),
                    (byte)(96 + ((x + y) & 159))));
            }
        }
        return image;
    }
}
