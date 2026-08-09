using System;
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

    private static int FindPngChunk(byte[] bytes, string expectedType) {
        int offset = 8;
        while (offset + 12 <= bytes.Length) {
            int length = ReadBigEndianInt32(bytes, offset);
            string type = System.Text.Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == expectedType) return offset;
            offset += 12 + length;
        }
        throw new InvalidDataException("PNG chunk was not found.");
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
