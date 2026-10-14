using System.IO.Compression;

namespace OfficeIMO.Archives.Tests;

public sealed partial class ZipExtractionTests {
    [Fact]
    public void MemoryReadsReturnVerifiedSelectedContentIndependentOfTheSource() {
        byte[] zip = Zip(("z.txt", "last", CompressionLevel.NoCompression), ("Templates/", null, CompressionLevel.Optimal),
            ("Templates/a.xml", "<policy />", CompressionLevel.Optimal));
        using var input = new MemoryStream(zip, writable: false);
        var files = ArchiveReader.ReadZip(input, new() { IncludeFile = path => path.EndsWith(".xml", StringComparison.Ordinal) });
        Assert.True(input.CanRead);
        Assert.Equal(input.Length, input.Position);
        var file = Assert.Single(files);
        Array.Clear(zip, 0, zip.Length);
        Assert.Equal("Templates/a.xml", file.RelativePath);
        Assert.Equal("<policy />", Encoding.UTF8.GetString(file.Data));
        Assert.Equal(file.Data.LongLength, file.Length);
    }

    [Fact]
    public void NonseekableReadsConsumeTheCurrentSourceWithoutClosingIt() {
        using var input = new NonseekableZip(Zip(("b", "B", CompressionLevel.Optimal), ("a", "A", CompressionLevel.NoCompression)));
        Assert.Equal(new[] { "a", "b" }, ArchiveReader.ReadZip(input).Select(file => file.RelativePath));
        Assert.True(input.CanRead);
        Assert.Equal(-1, input.ReadByte());
    }

    [Theory]
    [InlineData("input")]
    [InlineData("entries")]
    [InlineData("entry")]
    [InlineData("expanded")]
    public void InMemoryReadsEnforceFiniteLimitsBeforeReturningAnyResult(string limit) {
        byte[] zip = Zip(("a", "AAA", CompressionLevel.Optimal), ("b", "BBB", CompressionLevel.NoCompression));
        ArchiveReadOptions options = new();
        switch (limit) {
            case "input": options.MaxInputBytes = zip.Length - 1; break;
            case "entries": options.MaxEntryCount = 1; break;
            case "entry": options.MaxEntryBytes = 2; break;
            case "expanded": options.MaxExpandedBytes = 5; break;
        }
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(zip, options));
        using var stream = new NonseekableZip(zip);
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(stream, options));
    }

    [Fact]
    public void SelectedPayloadCorruptionFailsTheWholeMemoryRead() {
        byte[] zip = PatchHeaders(Zip(("a", "AAA", CompressionLevel.NoCompression)), crcXor: 1);
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(zip));
    }

    [Fact]
    public void ExcludedMembersStillUseTheStrictSharedMetadataAndPathRules() {
        var selectNone = new ArchiveReadOptions { IncludeFile = _ => false };
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(Zip(("../escape", "A", CompressionLevel.Optimal)), selectNone));
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(Zip(("a", "A", CompressionLevel.Optimal), ("A", "B", CompressionLevel.Optimal)), selectNone));
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadZip(PatchHeaders(Zip(("a", "A", CompressionLevel.Optimal)), flagsToSet: 1), selectNone));
    }

    [Fact]
    public void CancellationDuringSelectionReturnsNoPartialContent() {
        using var cancel = new CancellationTokenSource();
        byte[] zip = Zip(("a", "A", CompressionLevel.Optimal), ("b", "B", CompressionLevel.Optimal));
        Assert.Throws<OperationCanceledException>(() => ArchiveReader.ReadZip(zip,
            new() { IncludeFile = _ => { cancel.Cancel(); return true; } }, cancel.Token));
    }

    private sealed class NonseekableZip(byte[] bytes) : Stream {
        private readonly MemoryStream _inner = new(bytes, writable: false);
        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }
}
