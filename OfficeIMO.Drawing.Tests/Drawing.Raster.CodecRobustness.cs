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
    public void BoundedPngAndTiffDecodeObserveCancellationInsideValidationAndCodecWork() {
        var source = new OfficeRasterImage(4096, 1025, OfficeColor.FromRgba(24, 80, 160, 224));
        byte[][] encoded = {
            CreatePngWithLargeAncillaryPayload(),
            OfficeTiffCodec.Encode(source, new OfficeTiffEncodeOptions {
                Compression = OfficeTiffCompression.Lzw,
                Predictor = OfficeTiffPredictor.Horizontal
            })
        };

        foreach (byte[] bytes in encoded) {
            using var cancellation = new CancellationTokenSource();
            var options = new OfficeRasterDecodeOptions {
                CancellationToken = cancellation.Token
            };
            var cancelThread = new Thread(() => {
                Thread.Sleep(10);
                cancellation.Cancel();
            }) { IsBackground = true };
            cancelThread.Start();

            try {
                Assert.Throws<OperationCanceledException>(() =>
                    OfficeRasterImageDecoder.TryDecode(bytes, options, out _, out _));
            } finally {
                Assert.True(cancelThread.Join(TimeSpan.FromSeconds(5)));
            }
        }
    }

    [Fact]
    public void BoundedPngIdentificationObservesCancellationDuringChunkValidation() {
        byte[] png = CreatePngWithLargeAncillaryPayload();
        using var cancellation = new CancellationTokenSource();
        var cancelThread = new Thread(() => {
            Thread.Sleep(10);
            cancellation.Cancel();
        }) { IsBackground = true };
        cancelThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                OfficeImageReader.TryIdentifyByContent(png, null, cancellation.Token, out _));
        } finally {
            Assert.True(cancelThread.Join(TimeSpan.FromSeconds(5)));
        }
    }

    [Fact]
    public void WideSingleRowBmpDecodeObservesCancellationInsidePixelLoops() {
        byte[] bmp = CreateWideBmp32(width: 8 * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        var cancelThread = new Thread(() => {
            Thread.Sleep(1);
            cancellation.Cancel();
        }) { IsBackground = true };
        cancelThread.Start();

        try {
            Assert.Throws<OperationCanceledException>(() =>
                OfficeBmpReader.TryDecode(bmp, cancellation.Token, out _));
        } finally {
            Assert.True(cancelThread.Join(TimeSpan.FromSeconds(5)));
        }
    }

    private static byte[] CreatePngWithLargeAncillaryPayload() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int iendOffset = png.Length - 12;
        const int payloadLength = 16 * 1024 * 1024;
        var chunk = new byte[payloadLength + 12];
        WriteBigEndianInt32(chunk, 0, payloadLength);
        chunk[4] = (byte)'v';
        chunk[5] = (byte)'p';
        chunk[6] = (byte)'A';
        chunk[7] = (byte)'g';
        uint crc = ComputeCrc(chunk, 4, payloadLength + 4);
        WriteBigEndianInt32(chunk, payloadLength + 8, unchecked((int)crc));

        var result = new byte[png.Length + chunk.Length];
        Buffer.BlockCopy(png, 0, result, 0, iendOffset);
        Buffer.BlockCopy(chunk, 0, result, iendOffset, chunk.Length);
        Buffer.BlockCopy(png, iendOffset, result, iendOffset + chunk.Length, 12);
        return result;
    }

    private static byte[] CreateWideBmp32(int width) {
        const int pixelOffset = 54;
        int pixelBytes = checked(width * 4);
        var bmp = new byte[checked(pixelOffset + pixelBytes)];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        WriteLittleEndianInt32(bmp, 2, bmp.Length);
        WriteLittleEndianInt32(bmp, 10, pixelOffset);
        WriteLittleEndianInt32(bmp, 14, 40);
        WriteLittleEndianInt32(bmp, 18, width);
        WriteLittleEndianInt32(bmp, 22, 1);
        bmp[26] = 1;
        bmp[28] = 32;
        return bmp;
    }

    private static void WriteLittleEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static uint ComputeCrc(byte[] bytes, int offset, int count) {
        uint crc = 0xFFFFFFFFU;
        for (int index = 0; index < count; index++) {
            crc ^= bytes[offset + index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
            }
        }
        return crc ^ 0xFFFFFFFFU;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
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
