using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void PngReaderAndExportResultRejectInvalidChunkCrc() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        png[29] ^= 0x01;

        Assert.False(OfficePngReader.TryGetFrameCount(png, out _));
        Assert.False(OfficePngReader.TryDecode(png, out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png));
    }

    [Fact]
    public void PngReaderAndExportResultRejectInvalidZlibChecksumWithValidChunkCrc() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int idatOffset = FindPngChunk(png, "IDAT");
        int length = ReadBigEndianInt32(png, idatOffset);
        png[idatOffset + 8 + length - 1] ^= 0x01;
        WritePngChunkCrc(png, idatOffset, length);

        Assert.True(OfficePngReader.TryGetFrameCount(png, out _));
        Assert.False(OfficePngReader.TryDecode(png, out _));
        Assert.Throws<ArgumentException>(() =>
            new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png));
    }

    [Fact]
    public void ExportResultAcceptsStructurallyValidAdam7PngWithoutManagedRasterDecode() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        png[28] = 1;
        WritePngChunkCrc(png, 8, 13);

        Assert.True(OfficePngReader.TryGetFrameCount(png, out int frameCount));
        Assert.Equal(1, frameCount);
        Assert.False(OfficePngReader.TryDecode(png, out _));
        var result = new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png);
        Assert.Equal(png, result.Bytes);
    }

    [Fact]
    public void PngContainerRejectsUnknownCriticalAndMisplacedPaletteChunks() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] unknownCritical = InsertPngChunkBefore(png, "IDAT", "ABCD", Array.Empty<byte>());
        byte[] misplacedPalette = InsertPngChunkBefore(png, "IEND", "PLTE", new byte[] { 0, 0, 0 });

        Assert.False(OfficePngReader.TryGetFrameCount(unknownCritical, out _));
        Assert.False(OfficePngReader.TryGetFrameCount(misplacedPalette, out _));
    }

    [Fact]
    public void PngContainerRejectsNonContiguousImageDataChunks() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] split = SplitPngImageDataWithAncillaryChunk(png);

        Assert.False(OfficePngReader.TryGetFrameCount(split, out _));
        Assert.False(OfficeImageReader.TryValidateContent(split, "split.png", out _));
    }

    [Fact]
    public void CompleteContentValidationRejectsTruncatedGifCorruptPngAndMarkerOnlyJpeg() {
        byte[] truncatedGif = { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 1, 0, 1, 0, 0, 0, 0 };
        byte[] markerOnlyJpeg = {
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0xFF, 0xD9
        };
        byte[] corruptPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        int idatOffset = FindPngChunk(corruptPng, "IDAT");
        int length = ReadBigEndianInt32(corruptPng, idatOffset);
        corruptPng[idatOffset + 8 + length - 1] ^= 0x01;
        WritePngChunkCrc(corruptPng, idatOffset, length);

        Assert.True(OfficeImageReader.TryIdentifyByContent(truncatedGif, "truncated.gif", out _));
        Assert.False(OfficeImageReader.TryValidateContent(truncatedGif, "truncated.gif", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(corruptPng, "corrupt.png", out _));
        Assert.False(OfficeImageReader.TryValidateContent(corruptPng, "corrupt.png", out _));
        Assert.True(OfficeImageReader.TryIdentifyByContent(markerOnlyJpeg, "marker-only.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(markerOnlyJpeg, "marker-only.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationChecksEveryApngFrameAndJpegScan() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);
        Assert.True(OfficeImageReader.TryValidateContent(apng, "animated.png", out _));
        Assert.True(OfficePngReader.TryGetFrameCount(apng, out int frameCount));
        Assert.Equal(2, frameCount);

        byte[] corruptApng = (byte[])apng.Clone();
        int frameDataOffset = FindPngChunk(corruptApng, "fdAT");
        int frameDataLength = ReadBigEndianInt32(corruptApng, frameDataOffset);
        corruptApng[frameDataOffset + 8 + frameDataLength - 1] ^= 0x01;
        WritePngChunkCrc(corruptApng, frameDataOffset, frameDataLength);
        Assert.False(OfficeImageReader.TryValidateContent(corruptApng, "animated.png", out _));

        byte[] jpegWithEmptyFinalScan = {
            0xFF, 0xD8,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01, 0x01, 0x11, 0x00,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0x01,
            0xFF, 0xDA, 0x00, 0x08, 0x01, 0x01, 0x00, 0x00, 0x3F, 0x00,
            0xFF, 0xD9
        };
        Assert.True(OfficeImageReader.TryIdentifyByContent(jpegWithEmptyFinalScan, "progressive.jpg", out _));
        Assert.False(OfficeImageReader.TryValidateContent(jpegWithEmptyFinalScan, "progressive.jpg", out _));
    }

    [Fact]
    public void CompleteContentValidationRejectsInvalidApngSequenceCountAndFrameBounds() {
        byte[] staticPng = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        byte[] apng = CreateTwoFrameApng(staticPng);

        byte[] invalidSequence = (byte[])apng.Clone();
        int frameDataOffset = FindPngChunk(invalidSequence, "fdAT");
        WriteBigEndianInt32(invalidSequence, frameDataOffset + 8, 7);
        WritePngChunkCrc(invalidSequence, frameDataOffset, ReadBigEndianInt32(invalidSequence, frameDataOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidSequence, "sequence.png", out _));

        byte[] invalidCount = (byte[])apng.Clone();
        int animationControlOffset = FindPngChunk(invalidCount, "acTL");
        WriteBigEndianInt32(invalidCount, animationControlOffset + 8, 3);
        WritePngChunkCrc(invalidCount, animationControlOffset, ReadBigEndianInt32(invalidCount, animationControlOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidCount, "count.png", out _));

        byte[] invalidBounds = (byte[])apng.Clone();
        int firstFrameControlOffset = FindPngChunk(invalidBounds, "fcTL");
        int secondFrameControlOffset = FindPngChunk(
            invalidBounds,
            "fcTL",
            firstFrameControlOffset + 12 + ReadBigEndianInt32(invalidBounds, firstFrameControlOffset));
        WriteBigEndianInt32(invalidBounds, secondFrameControlOffset + 12, 2);
        WritePngChunkCrc(invalidBounds, secondFrameControlOffset, ReadBigEndianInt32(invalidBounds, secondFrameControlOffset));
        Assert.False(OfficeImageReader.TryValidateContent(invalidBounds, "bounds.png", out _));
    }

    [Fact]
    public async Task GuardedAsyncConsumerSerializesConcurrentAdmissionAndSequenceAssignment() {
        const int maximum = 300;
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var results = new List<OfficeImageExportResult>();
        var options = new OfficeImageExportOptions { MaximumOutputCount = maximum };

        await Assert.ThrowsAsync<OfficeImageExportBatchLimitException>(() =>
            OfficeImageExportBatchProcessor.RunAsync(
                options,
                async (accept, token) => await Task.WhenAll(
                    Enumerable.Range(0, 500).Select(index => accept(
                        new OfficeImageExportResult(
                            OfficeImageExportFormat.Png,
                            1,
                            1,
                            png,
                            name: index.ToString()),
                        token))),
                (result, _) => {
                    results.Add(result);
                    return Task.CompletedTask;
                }));

        Assert.Equal(maximum, results.Count);
        Assert.Equal(Enumerable.Range(0, maximum), results.Select(result => result.SequenceIndex!.Value));
    }

    [Fact]
    public async Task GuardedAsyncConsumerAllowsConsumerReentry() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var observed = new ConcurrentQueue<int>();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 2 },
            async (result, token) => {
                observed.Enqueue(result.SequenceIndex!.Value);
                if (result.Name == "outer") {
                    await accept!(new OfficeImageExportResult(
                        OfficeImageExportFormat.Png,
                        1,
                        1,
                        png,
                        name: "inner"), token);
                }
            },
            cancellation.Token);

        await accept(new OfficeImageExportResult(
            OfficeImageExportFormat.Png,
            1,
            1,
            png,
            name: "outer"), cancellation.Token);

        Assert.Equal(new[] { 0, 1 }, observed);
    }

    [Fact]
    public async Task GuardedAsyncConsumerRejectsForkedReentryWithoutConcurrentAdmission() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var observed = new ConcurrentQueue<int>();
        OfficeImageExportAsyncConsumer? accept = null;
        accept = OfficeImageExportBatchProcessor.CreateGuardedAsyncConsumer(
            new OfficeImageExportOptions { MaximumOutputCount = 3 },
            async (result, token) => {
                observed.Enqueue(result.SequenceIndex!.Value);
                if (result.Name == "outer") {
                    await Task.WhenAll(
                        Task.Run(() => accept!(CreateResult("fork-1", png), token), token),
                        Task.Run(() => accept!(CreateResult("fork-2", png), token), token));
                }
            });

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            accept(CreateResult("outer", png), default));
        Assert.Equal(new[] { 0 }, observed);
    }

    [Fact]
    public void GuardedConsumerSerializesConcurrentAdmissionAndSequenceAssignment() {
        const int maximum = 300;
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var results = new List<OfficeImageExportResult>();
        var options = new OfficeImageExportOptions { MaximumOutputCount = maximum };
        OfficeImageExportConsumer accept = OfficeImageExportBatchProcessor.CreateGuardedConsumer(
            options,
            result => results.Add(result));
        int rejected = 0;

        System.Threading.Tasks.Parallel.For(0, 500, index => {
            try {
                accept(new OfficeImageExportResult(
                    OfficeImageExportFormat.Png,
                    1,
                    1,
                    png,
                    name: index.ToString()));
            } catch (OfficeImageExportBatchLimitException) {
                Interlocked.Increment(ref rejected);
            }
        });

        Assert.Equal(maximum, results.Count);
        Assert.Equal(200, rejected);
        Assert.Equal(Enumerable.Range(0, maximum), results.Select(result => result.SequenceIndex!.Value));
    }

    [Fact]
    public void EffectiveScaleUsesTargetDpiWithoutRequiringValidationSideEffects() {
        var options = new OfficeImageExportOptions {
            Scale = 1D,
            TargetDpi = 192D
        };

        Assert.Equal(2D, options.GetEffectiveScale(100D, 100D));
        Assert.Equal(1D, options.Scale);
    }

    [Fact]
    public void ImageResultRejectsUndefinedFileConflictPolicyBeforeWriting() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var result = new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1, png);
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-" + Guid.NewGuid().ToString("N") + ".png");

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            result.Save(path, (OfficeImageExportFileConflictPolicy)999));
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void StreamIdentificationRejectsOversizedSeekablePayloadBeforeReading() {
        using var stream = new DeclaredLengthStream(128L * 1024L * 1024L + 1L);

        Assert.False(OfficeImageReader.TryIdentifyByContent(stream, "oversized.png", out _));
        Assert.Equal(0, stream.ReadCount);
        Assert.Equal(0L, stream.Position);
    }

    private static int FindPngChunk(byte[] bytes, string expectedType, int offset = 8) {
        while (offset + 12 <= bytes.Length) {
            int length = ReadBigEndianInt32(bytes, offset);
            string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == expectedType) return offset;
            offset += 12 + length;
        }
        throw new InvalidDataException("PNG chunk was not found.");
    }

    private static OfficeImageExportResult CreateResult(string name, byte[] png) =>
        new(OfficeImageExportFormat.Png, 1, 1, png, name: name);

    private static byte[] CreateTwoFrameApng(byte[] png) {
        int idatOffset = FindPngChunk(png, "IDAT");
        int idatLength = ReadBigEndianInt32(png, idatOffset);
        int idatEnd = idatOffset + 12 + idatLength;
        byte[] animationControl = new byte[8];
        WriteBigEndianInt32(animationControl, 0, 2);
        byte[] firstFrameControl = CreateFrameControl(sequence: 0);
        byte[] secondFrameControl = CreateFrameControl(sequence: 1);
        byte[] secondFrameData = new byte[idatLength + 4];
        WriteBigEndianInt32(secondFrameData, 0, 2);
        Buffer.BlockCopy(png, idatOffset + 8, secondFrameData, 4, idatLength);
        byte[] prefix = CreatePngChunk("acTL", animationControl)
            .Concat(CreatePngChunk("fcTL", firstFrameControl))
            .ToArray();
        byte[] suffix = CreatePngChunk("fcTL", secondFrameControl)
            .Concat(CreatePngChunk("fdAT", secondFrameData))
            .ToArray();
        byte[] result = new byte[png.Length + prefix.Length + suffix.Length];
        Buffer.BlockCopy(png, 0, result, 0, idatOffset);
        Buffer.BlockCopy(prefix, 0, result, idatOffset, prefix.Length);
        Buffer.BlockCopy(png, idatOffset, result, idatOffset + prefix.Length, idatEnd - idatOffset);
        Buffer.BlockCopy(suffix, 0, result, idatEnd + prefix.Length, suffix.Length);
        Buffer.BlockCopy(png, idatEnd, result, idatEnd + prefix.Length + suffix.Length, png.Length - idatEnd);
        return result;
    }

    private static byte[] CreateFrameControl(int sequence) {
        byte[] data = new byte[26];
        WriteBigEndianInt32(data, 0, sequence);
        WriteBigEndianInt32(data, 4, 1);
        WriteBigEndianInt32(data, 8, 1);
        data[21] = 1;
        return data;
    }

    private static void WritePngChunkCrc(byte[] bytes, int chunkOffset, int length) {
        uint crc = 0xFFFFFFFFU;
        for (int index = chunkOffset + 4; index < chunkOffset + 8 + length; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
            }
        }
        crc ^= 0xFFFFFFFFU;
        int offset = chunkOffset + 8 + length;
        bytes[offset] = (byte)(crc >> 24);
        bytes[offset + 1] = (byte)(crc >> 16);
        bytes[offset + 2] = (byte)(crc >> 8);
        bytes[offset + 3] = (byte)crc;
    }

    private static byte[] InsertPngChunkBefore(byte[] png, string beforeType, string newType, byte[] data) {
        int beforeOffset = FindPngChunk(png, beforeType);
        byte[] chunk = CreatePngChunk(newType, data);
        byte[] result = new byte[png.Length + chunk.Length];
        Buffer.BlockCopy(png, 0, result, 0, beforeOffset);
        Buffer.BlockCopy(chunk, 0, result, beforeOffset, chunk.Length);
        Buffer.BlockCopy(png, beforeOffset, result, beforeOffset + chunk.Length, png.Length - beforeOffset);
        return result;
    }

    private static byte[] SplitPngImageDataWithAncillaryChunk(byte[] png) {
        int idatOffset = FindPngChunk(png, "IDAT");
        int length = ReadBigEndianInt32(png, idatOffset);
        int firstLength = length / 2;
        int secondLength = length - firstLength;
        byte[] first = new byte[firstLength];
        byte[] second = new byte[secondLength];
        Buffer.BlockCopy(png, idatOffset + 8, first, 0, firstLength);
        Buffer.BlockCopy(png, idatOffset + 8 + firstLength, second, 0, secondLength);
        byte[] firstChunk = CreatePngChunk("IDAT", first);
        byte[] ancillary = CreatePngChunk("vpAg", Array.Empty<byte>());
        byte[] secondChunk = CreatePngChunk("IDAT", second);
        int originalChunkEnd = idatOffset + 12 + length;
        byte[] result = new byte[png.Length - (12 + length) + firstChunk.Length + ancillary.Length + secondChunk.Length];
        int offset = 0;
        Buffer.BlockCopy(png, 0, result, offset, idatOffset);
        offset += idatOffset;
        Buffer.BlockCopy(firstChunk, 0, result, offset, firstChunk.Length);
        offset += firstChunk.Length;
        Buffer.BlockCopy(ancillary, 0, result, offset, ancillary.Length);
        offset += ancillary.Length;
        Buffer.BlockCopy(secondChunk, 0, result, offset, secondChunk.Length);
        offset += secondChunk.Length;
        Buffer.BlockCopy(png, originalChunkEnd, result, offset, png.Length - originalChunkEnd);
        return result;
    }

    private static byte[] CreatePngChunk(string type, byte[] data) {
        byte[] chunk = new byte[12 + data.Length];
        WriteBigEndianInt32(chunk, 0, data.Length);
        System.Text.Encoding.ASCII.GetBytes(type, 0, 4, chunk, 4);
        Buffer.BlockCopy(data, 0, chunk, 8, data.Length);
        WritePngChunkCrc(chunk, 0, data.Length);
        return chunk;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private sealed class DeclaredLengthStream : Stream {
        private long _position;

        internal DeclaredLengthStream(long length) => Length = length;

        internal int ReadCount { get; private set; }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length { get; }
        public override long Position { get => _position; set => _position = value; }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { ReadCount++; return 0; }
        public override long Seek(long offset, SeekOrigin origin) => _position = offset;
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
