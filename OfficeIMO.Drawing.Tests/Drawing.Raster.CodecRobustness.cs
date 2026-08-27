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
            Thread cancelThread = StartCancellationThread(cancellation, delayMilliseconds: 10);

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
        Thread cancelThread = StartCancellationThread(cancellation, delayMilliseconds: 1);

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
        Thread cancelThread = StartCancellationThread(cancellation, delayMilliseconds: 1);

        try {
            Assert.Throws<OperationCanceledException>(() =>
                OfficeBmpReader.TryDecode(bmp, cancellation.Token, out _));
        } finally {
            Assert.True(cancelThread.Join(TimeSpan.FromSeconds(5)));
        }
    }

    [Fact]
    public void TiffInspectionPassesCancellationIntoStructureValidation() {
        byte[] tiff = CreateTiffWithMaximumEntryInventory();
        using var cancellation = new CancellationTokenSource();
        var options = new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token };
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            OfficeTiffCodec.TryInspectPages(tiff, options, out _));
    }

    [Fact]
    public void WideSingleRowTiffDecodeObservesCancellationInsidePixelLoops() {
        byte[] tiff = CreateWideGrayscaleTiff(width: 8 * 1024 * 1024);
        using var cancellation = new CancellationTokenSource();
        var options = new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token };
        Thread cancelThread = StartCancellationThread(cancellation, delayMilliseconds: 10);

        try {
            Assert.Throws<OperationCanceledException>(() =>
                OfficeTiffCodec.TryDecodePage(tiff, 0, options, out _));
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

    private static Thread StartCancellationThread(
        CancellationTokenSource cancellation,
        int delayMilliseconds) {
        using var ready = new ManualResetEventSlim();
        var thread = new Thread(() => {
            ready.Set();
            Thread.Sleep(delayMilliseconds);
            cancellation.Cancel();
        }) { IsBackground = true };
        thread.Start();
        Assert.True(ready.Wait(TimeSpan.FromSeconds(5)));
        return thread;
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

    private static byte[] CreateTiffWithMaximumEntryInventory() {
        const int entryCount = ushort.MaxValue;
        const int ifdOffset = 8;
        var tiff = new byte[ifdOffset + 2 + entryCount * 12 + 4];
        tiff[0] = (byte)'I';
        tiff[1] = (byte)'I';
        tiff[2] = 42;
        WriteLittleEndianInt32(tiff, 4, ifdOffset);
        WriteLittleEndianUInt16(tiff, ifdOffset, entryCount);
        int entryOffset = ifdOffset + 2;
        for (int tag = 0; tag < entryCount; tag++, entryOffset += 12) {
            WriteLittleEndianUInt16(tiff, entryOffset, tag);
            WriteLittleEndianUInt16(tiff, entryOffset + 2, tag is 256 or 257 ? 4 : 1);
            WriteLittleEndianInt32(tiff, entryOffset + 4, 1);
            WriteLittleEndianInt32(tiff, entryOffset + 8, tag is 256 or 257 ? 1 : 0);
        }
        return tiff;
    }

    private static byte[] CreateWideGrayscaleTiff(int width) {
        const int entryCount = 9;
        const int ifdOffset = 8;
        int pixelOffset = ifdOffset + 2 + entryCount * 12 + 4;
        var tiff = new byte[checked(pixelOffset + width)];
        tiff[0] = (byte)'I';
        tiff[1] = (byte)'I';
        tiff[2] = 42;
        WriteLittleEndianInt32(tiff, 4, ifdOffset);
        WriteLittleEndianUInt16(tiff, ifdOffset, entryCount);
        int entryOffset = ifdOffset + 2;
        WriteTiffEntry(tiff, ref entryOffset, 256, 4, width);
        WriteTiffEntry(tiff, ref entryOffset, 257, 4, 1);
        WriteTiffEntry(tiff, ref entryOffset, 258, 3, 8);
        WriteTiffEntry(tiff, ref entryOffset, 259, 3, 1);
        WriteTiffEntry(tiff, ref entryOffset, 262, 3, 1);
        WriteTiffEntry(tiff, ref entryOffset, 273, 4, pixelOffset);
        WriteTiffEntry(tiff, ref entryOffset, 277, 3, 1);
        WriteTiffEntry(tiff, ref entryOffset, 278, 4, 1);
        WriteTiffEntry(tiff, ref entryOffset, 279, 4, width);
        return tiff;
    }

    private static void WriteTiffEntry(
        byte[] bytes,
        ref int offset,
        int tag,
        int type,
        int value) {
        WriteLittleEndianUInt16(bytes, offset, tag);
        WriteLittleEndianUInt16(bytes, offset + 2, type);
        WriteLittleEndianInt32(bytes, offset + 4, 1);
        WriteLittleEndianInt32(bytes, offset + 8, value);
        offset += 12;
    }

    private static void WriteLittleEndianUInt16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
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
