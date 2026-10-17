using System.IO.Compression;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveZipInspectionTests {
    [Fact]
    public void TraversalRetainsThePlatformDosTimeInterpretationAcrossInputOverloads() {
        using var source = new MemoryStream();
        var wallTime = new DateTime(2024, 1, 2, 3, 4, 6, DateTimeKind.Unspecified);
        using (var writer = new ZipArchive(source, ZipArchiveMode.Create, true)) {
            writer.CreateEntry("member").LastWriteTime = new DateTimeOffset(wallTime, TimeSpan.Zero);
        }
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, true);
        DateTime expected = archive.Entries[0].LastWriteTime.UtcDateTime;
        using var inspection = ArchiveInspector.Open(source);
        Assert.Equal(wallTime, inspection.Entries[0].LastWriteTime!.Value.DateTime);
        Assert.Equal(TimeSpan.Zero, inspection.Entries[0].LastWriteTime!.Value.Offset);
        Assert.Equal(expected, Assert.Single(ArchiveReader.TraverseZip(archive).Entries).LastWriteUtc);
        Assert.Equal(expected, Assert.Single(ArchiveReader.TraverseZip(inspection).Entries).LastWriteUtc);
        Assert.Equal(expected, Assert.Single(ArchiveReader.TraverseZip(source).Entries).LastWriteUtc);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(1250)]
    public void TraversalRejectsMalformedDeclaredUtf8WithOrWithoutALegacyPolicy(int? codePage) {
        byte[] bytes = Zip(("ab", "content"));
        int central = Central(bytes);
        Set16(bytes, 6, 0x0800); Set16(bytes, central + 8, 0x0800);
        bytes[30] = bytes[central + 46] = 0xc3;
        bytes[31] = bytes[central + 47] = 0x28;
        var options = new ArchiveTraversalOptions { ZipLegacyCodePage = codePage };
        using var source = new MemoryStream(bytes);
        Assert.Throws<InvalidDataException>(() => ArchiveReader.TraverseZip(source, options));
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveReader.TraverseZip(fixture.Package(bytes), options));
    }

    [Fact]
    public void NonseekableInspectionConsumesOnlyTheRemainingInputAndLeavesItOpen() {
        using var input = new NonseekableInput(Zip(("member", "complete")));
        using var inspection = ArchiveInspector.Open(input);
        using var member = inspection.OpenRead(0);
        Assert.Equal("complete", Text(member));
        Assert.True(input.CanRead);
        Assert.Equal(-1, input.ReadByte());
    }
    [Fact]
    public void PassiveZipInspectionPreservesUnsafeDuplicateIdentitiesAndDetachedVerifiedBatches() {
        byte[] bytes = Zip(("../same.txt", "first"), ("../same.txt", "second"));
        using var input = new MemoryStream(bytes);
        var inspection = ArchiveInspector.Open(input);
        Assert.Equal(new[] { "../same.txt", "../same.txt" }, inspection.Entries.Select(e => e.RawName));
        using var batch = inspection.ReadMembers(new[] { 1, 0 }, 11);
        inspection.Dispose();
        Assert.Equal(new[] { "second", "first" }, batch.Members.Select(m => Text(m.Content)));
        Assert.Equal(11, batch.SelectedBytes);
        Assert.Throws<ArgumentException>(() => {
            using var other = ArchiveInspector.Open(new MemoryStream(bytes));
            other.ReadMembers(new[] { 0, 0 }, 12);
        });
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(bytes), fixture.Destination));
    }

    [Fact]
    public void VerificationReportsCorruptionAndUnsupportedMembersWithoutReturningPrefixes() {
        byte[] bytes = Zip(("bad.txt", "hello"), ("unsupported.txt", "world"));
        int central = Central(bytes);
        bytes[Payload(bytes)] ^= 1;
        int second = central + 46 + BitConverter.ToUInt16(bytes, central + 28) + BitConverter.ToUInt16(bytes, central + 30) + BitConverter.ToUInt16(bytes, central + 32);
        Set16(bytes, second + 10, 99);
        using var inspection = ArchiveInspector.Open(new MemoryStream(bytes));
        Assert.False(inspection.Entries[1].CanRead);
        var result = inspection.Verify();
        Assert.Equal(new[] { ArchiveVerificationStatus.Invalid, ArchiveVerificationStatus.Unsupported }, result.Members.Select(m => m.Status));
        Assert.Equal("ARCHIVE_INTEGRITY", result.Members[0].Code);
        Assert.Equal("ZIP_COMPRESSION", result.Members[1].Code);
        Assert.Throws<InvalidDataException>(() => inspection.ReadMembers(new[] { 0 }, 5));
        Assert.True(result.ResourceUsage.DecompressionWorkBytes > 0);
    }

    [Fact]
    public void ExplicitLegacyPolishNamesAreSharedByInspectionMemoryReadsAndExtraction() {
        // CP1250 bytes for Łódź.txt, authored independently of the decoder's implementation.
        byte[] name = { 0xa3, 0xf3, (byte)'d', 0x9f, (byte)'.', (byte)'t', (byte)'x', (byte)'t' };
        byte[] bytes = Zip(("12345678", "content"));
        name.CopyTo(bytes, 30); name.CopyTo(bytes, Central(bytes) + 46);
        Assert.Throws<NotSupportedException>(() => ArchiveInspector.Open(new MemoryStream(bytes)));
        using var inspection = ArchiveInspector.Open(new MemoryStream(bytes), new() { ZipLegacyCodePage = 1250 });
        Assert.Equal("Łódź.txt", inspection.Entries[0].RawName);
        using var content = inspection.OpenRead(0);
        Assert.Equal("content", Text(content));
        Assert.Equal("Łódź.txt", Assert.Single(ArchiveReader.ReadZip(bytes, new() { ZipLegacyCodePage = 1250 })).RelativePath);
        using var fixture = new ArchiveFixtures();
        var extracted = ArchiveExtractor.Extract(fixture.Package(bytes), fixture.Destination, new() { ZipLegacyCodePage = 1250 });
        Assert.Equal("Łódź.txt", Assert.Single(extracted.Files).RelativePath);
        Assert.Equal("content", File.ReadAllText(Path.Combine(fixture.Destination, "Łódź.txt")));
    }

    [Fact]
    public void Utf8DeclarationOverridesTheCallerLegacyEncoding() {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true, Encoding.UTF8)) {
            using var writer = new StreamWriter(zip.CreateEntry("zażółć.txt").Open()); writer.Write("valid");
        }
        using var inspection = ArchiveInspector.Open(new MemoryStream(output.ToArray()), new() { ZipLegacyCodePage = 1252 });
        Assert.Equal("zażółć.txt", inspection.Entries[0].RawName);
        Assert.Equal(ArchiveVerificationStatus.Verified, Assert.Single(inspection.Verify().Members).Status);
    }

    internal static byte[] Zip(params (string Name, string Content)[] entries) {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var entry in entries) {
            using var writer = new StreamWriter(zip.CreateEntry(entry.Name, CompressionLevel.NoCompression).Open(), new UTF8Encoding(false));
            writer.Write(entry.Content);
        }
        return output.ToArray();
    }
    internal static int Central(byte[] bytes) => (int)BitConverter.ToUInt32(bytes, bytes.Length - 6);
    internal static int Payload(byte[] bytes) => 30 + BitConverter.ToUInt16(bytes, 26) + BitConverter.ToUInt16(bytes, 28);
    internal static void Set16(byte[] bytes, int offset, ushort value) => BitConverter.GetBytes(value).CopyTo(bytes, offset);
    internal static string Text(Stream stream) { using var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, true); return reader.ReadToEnd(); }

    private sealed class NonseekableInput(byte[] bytes) : Stream {
        private readonly MemoryStream _source = new(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => _source.Read(buffer, offset, count);
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) _source.Dispose(); base.Dispose(disposing); }
    }
}
