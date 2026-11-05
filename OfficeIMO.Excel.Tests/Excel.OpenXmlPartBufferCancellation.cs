#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.IO.Compression;
using System.Threading;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void PackagePartReaderDoesNotReadCompressedTailAfterCancellation() {
        byte[] expected = Enumerable.Repeat((byte)'x', 2 * 1024 * 1024 + 256).ToArray();
        var fixture = CreatePartWithEmptyDeflateTail(expected);
        using var cancellation = new CancellationTokenSource();
        using var input = new TailCancellationInput(fixture.Package, fixture.TailStart, cancellation);
        using var parts = OpenXmlPackagePartBufferReader.TryOpen(input)!;
        Assert.NotNull(parts);
        input.Armed = true;
        byte[]? actual = null;
        int length = 0;
        try {
            Exception? error = Record.Exception(() => {
                Assert.True(parts.TryRead("large.bin", expected.Length, cancellation.Token, out actual, out length));
            });
            Assert.True(error == null || error is OperationCanceledException,
                error?.ToString() ?? "The part neither completed nor observed cancellation.");
            if (error == null) {
                Assert.Equal(expected.Length, length);
                Assert.True(expected.AsSpan().SequenceEqual(actual.AsSpan(0, length)));
            } else {
                Assert.True(cancellation.IsCancellationRequested);
            }
            Assert.Equal(0, input.ReadsAfterCancellation);
        } finally {
            if (actual != null) OpenXmlPartBufferPool.Return(actual);
        }
    }

    // Stored DEFLATE blocks allow a valid long compressed tail without any further output.
    private static (byte[] Package, int TailStart) CreatePartWithEmptyDeflateTail(byte[] content) {
        using var original = new MemoryStream();
        using (var zip = new ZipArchive(original, ZipArchiveMode.Create, leaveOpen: true)) {
            using var entry = zip.CreateEntry("large.bin", CompressionLevel.NoCompression).Open();
            entry.Write(content);
        }
        byte[] bytes = original.ToArray();
        int oldDirectory = (int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(bytes.Length - 6, 4));
        int header = 30 + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(26, 2))
            + BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(28, 2));
        using var compressed = new MemoryStream();
        for (int offset = 0; offset < content.Length;) {
            int count = Math.Min(ushort.MaxValue, content.Length - offset);
            compressed.WriteByte(0);
            compressed.WriteByte((byte)count); compressed.WriteByte((byte)(count >> 8));
            ushort inverse = unchecked((ushort)~count);
            compressed.WriteByte((byte)inverse); compressed.WriteByte((byte)(inverse >> 8));
            compressed.Write(content, offset, count);
            offset += count;
        }
        int tailStart = header + (int)compressed.Length;
        byte[] emptyBlock = { 0, 0, 0, 255, 255 };
        for (int index = 0; index < 400_000; index++) compressed.Write(emptyBlock);
        compressed.Write(new byte[] { 1, 0, 0, 255, 255 });
        byte[] raw = compressed.ToArray();
        int directory = header + raw.Length;
        byte[] package = new byte[directory + bytes.Length - oldDirectory];
        bytes.AsSpan(0, header).CopyTo(package);
        raw.CopyTo(package, header);
        bytes.AsSpan(oldDirectory).CopyTo(package.AsSpan(directory));
        BinaryPrimitives.WriteUInt16LittleEndian(package.AsSpan(8, 2), 8);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(18, 4), (uint)raw.Length);
        BinaryPrimitives.WriteUInt16LittleEndian(package.AsSpan(directory + 10, 2), 8);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(directory + 20, 4), (uint)raw.Length);
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(package.Length - 6, 4), (uint)directory);
        return (package, tailStart);
    }

    private sealed class TailCancellationInput : Stream {
        private readonly MemoryStream _input;
        private readonly int _tailStart;
        private readonly CancellationTokenSource _cancellation;
        internal TailCancellationInput(byte[] package, int tailStart, CancellationTokenSource cancellation) {
            _input = new MemoryStream(package, writable: false); _tailStart = tailStart; _cancellation = cancellation;
        }
        internal bool Armed { get; set; }
        internal int ReadsAfterCancellation { get; private set; }
        private void BeforeRead() {
            if (!Armed || Position < _tailStart) return;
            if (_cancellation.IsCancellationRequested) {
                ReadsAfterCancellation++;
                throw new InvalidOperationException("Compressed input was read again after cancellation.");
            }
            _cancellation.Cancel();
        }
        public override int Read(byte[] buffer, int offset, int count) { BeforeRead(); return _input.Read(buffer, offset, count); }
        public override int Read(Span<byte> buffer) { BeforeRead(); return _input.Read(buffer); }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _input.Length;
        public override long Position { get => _input.Position; set => _input.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => _input.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _input.Dispose(); base.Dispose(disposing); }
    }
}
#endif
