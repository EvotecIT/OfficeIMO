using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class DrawingHeifMetadataTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void KnownMemoryStreamBackingRemainsChargedDuringXmpDecoding(int streamShape) {
        byte[] bytes = CreateLargeHeifSource(exif: false, prefix: streamShape == 1 ? 37 : 0);
        if (streamShape == 0) {
            Assert.True(OfficeHeifMetadataReader.TryReadXmp(bytes, out _));
        }
        int prefix = streamShape == 1 ? 37 : 0;
        using var stream = new MemoryStream(bytes, prefix, bytes.Length - prefix,
            writable: false, publiclyVisible: streamShape != 2);
        Assert.False(OfficeHeifMetadataReader.TryReadXmp(stream, out string? packet));
        Assert.Null(packet);
        Assert.Equal(0L, stream.Position);
        Assert.True(stream.CanRead);
        Assert.Equal((byte)0, stream.ReadByte());
    }

    [Fact]
    public void KnownMemoryStreamBackingRemainsChargedThroughSharedExifParsing() {
        byte[] bytes = CreateLargeHeifSource(exif: true, prefix: 0);
        Assert.True(OfficeHeifMetadataReader.TryReadExifProfile(bytes, out _));
        using var stream = new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true);
        Assert.False(OfficeHeifMetadataReader.TryReadExifProfile(stream, out OfficeImageMetadata? profile));
        Assert.Null(profile);
        Assert.Equal(0L, stream.Position);
        Assert.True(stream.CanRead);
    }

    private static byte[] CreateLargeHeifSource(bool exif, int prefix) {
        const int mebibyte = 1024 * 1024;
        byte[] payload;
        if (exif) {
            var metadata = new OfficeImageMetadata();
            metadata.SetExifValue(OfficeExifTag.UserComment, new byte[4 * mebibyte]);
            payload = metadata.EncodeExifProfile()!;
        } else {
            payload = CoreHeifFixtures.CreateExifPayload("Original");
        }
        byte[] fixture = CoreHeifFixtures.CreateHeifMetadataSiblings(
            payload, exif ? "packet" : new string('X', 6 * mebibyte), 0, 0);
        var bytes = new byte[120 * mebibyte + prefix];
        Buffer.BlockCopy(fixture, 0, bytes, prefix, fixture.Length);
        int mediaDataOffset = prefix + FindAscii(fixture, "mdat") - 4;
        uint mediaDataLength = (uint)(bytes.Length - mediaDataOffset);
        bytes[mediaDataOffset] = (byte)(mediaDataLength >> 24);
        bytes[mediaDataOffset + 1] = (byte)(mediaDataLength >> 16);
        bytes[mediaDataOffset + 2] = (byte)(mediaDataLength >> 8);
        bytes[mediaDataOffset + 3] = (byte)mediaDataLength;
        return bytes;
    }

    [Fact]
    public void CancellationDuringStreamReadRestoresCallerPositionAndLeavesItOpen() {
        byte[] fixture = CoreHeifFixtures.CreateHeifMetadataSiblings(
            CoreHeifFixtures.CreateExifPayload("Original"), "packet", 0, 0);
        var bytes = new byte[fixture.Length + 3];
        Buffer.BlockCopy(fixture, 0, bytes, 3, fixture.Length);
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelDuringReadStream(bytes, cancellation);
        stream.Position = 3;
        Assert.Throws<OperationCanceledException>(() =>
            OfficeHeifMetadataReader.TryReadInfo(stream, out _, cancellation.Token));
        Assert.True(stream.ReadOccurred);
        Assert.Equal(3L, stream.Position);
        Assert.True(stream.CanRead);
    }

    private sealed class CancelDuringReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;
        internal bool ReadOccurred { get; private set; }
        internal CancelDuringReadStream(byte[] bytes, CancellationTokenSource cancellation) : base(bytes) {
            _cancellation = cancellation;
        }
        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            ReadOccurred = true;
            _cancellation.Cancel();
            return read;
        }
    }
}
