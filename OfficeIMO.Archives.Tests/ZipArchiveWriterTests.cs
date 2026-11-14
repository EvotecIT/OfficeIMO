using System.IO.Compression;
using System.Security.Cryptography;

namespace OfficeIMO.Archives.Tests;

public sealed class ZipArchiveWriterTests {
    internal static readonly byte[] Text = Encoding.UTF8.GetBytes(string.Concat(Enumerable.Repeat("Deterministic ZIP content. ", 400)));

    internal static ArchiveWriteEntry[] SampleEntries() => new[] {
        ArchiveWriteEntry.Directory("docs"),
        ArchiveWriteEntry.File("docs/readme.txt", Text),
        ArchiveWriteEntry.File("docs/zażółć-gęślą-jaźń.txt", Encoding.UTF8.GetBytes("Zażółć gęślą jaźń.\n")),
        ArchiveWriteEntry.File("empty.bin", Array.Empty<byte>()),
        ArchiveWriteEntry.File("data/random.bin", ArchiveDeflateEncoderTests.Sample("random").Take(70_000).ToArray())
    };

    [Fact]
    public void WrittenArchiveRoundTripsThroughOwnReaderExtractionAndTheRuntimeReader() {
        byte[] zip = Write(SampleEntries());

        var files = ArchiveReader.ReadZip(zip);
        Assert.Equal(new[] { "data/random.bin", "docs/readme.txt", "docs/zażółć-gęślą-jaźń.txt", "empty.bin" }, files.Select(file => file.RelativePath));
        Assert.Equal(Text, files.Single(file => file.RelativePath == "docs/readme.txt").Data);

        using (var runtime = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read)) {
            Assert.Equal(new[] { "docs/", "docs/readme.txt", "docs/zażółć-gęślą-jaźń.txt", "empty.bin", "data/random.bin" },
                runtime.Entries.Select(entry => entry.FullName));
            using var content = new MemoryStream();
            using (Stream stream = runtime.GetEntry("docs/readme.txt")!.Open()) stream.CopyTo(content);
            Assert.Equal(Text, content.ToArray());
            Assert.Equal(new DateTime(1980, 1, 1), runtime.GetEntry("empty.bin")!.LastWriteTime.DateTime);
        }

        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(zip), fixture.Destination);
        Assert.Equal(ArchiveFormat.Zip, result.Format);
        Assert.Equal(Text, File.ReadAllBytes(Path.Combine(fixture.Destination, "docs", "readme.txt")));
    }

    [Fact]
    public void DefaultOutputIsByteIdenticalAcrossRunsTargetsAndWriteRoutes() {
        byte[] first = Write(SampleEntries());
        Assert.Equal(first, Write(SampleEntries()));

        using var routed = new MemoryStream();
        using (var writer = new ZipArchiveWriter(routed, leaveOpen: true)) {
            writer.AddDirectory("docs/");
            using (Stream entry = writer.CreateEntry("docs/readme.txt"))
                for (int offset = 0; offset < Text.Length; offset += 13) entry.Write(Text, offset, Math.Min(13, Text.Length - offset));
            writer.AddFile("docs/zażółć-gęślą-jaźń.txt", new NonSeekableStream(Encoding.UTF8.GetBytes("Zażółć gęślą jaźń.\n")));
            writer.AddFile("empty.bin", new MemoryStream());
            writer.AddFile("data\\random.bin", ArchiveDeflateEncoderTests.Sample("random").Take(70_000).ToArray());
        }
        Assert.Equal(first, routed.ToArray());
        // The same digest is expected from every target framework and operating system.
        Assert.Equal("b7442d3ea36a3f6b", Hash(first));
    }

    [Fact]
    public void FixedTimestampAndSourceDateEpochClampingAreApplied() {
        var epoch = ArchiveTimestamp.ParseSourceDateEpoch("1700000000")!.Value;
        var older = new DateTimeOffset(2001, 2, 3, 4, 5, 6, TimeSpan.Zero);
        var newer = new DateTimeOffset(2030, 1, 1, 0, 0, 0, TimeSpan.Zero);
        byte[] zip = Write(new[] {
            ArchiveWriteEntry.File("older.txt", new byte[] { 1 }, older),
            ArchiveWriteEntry.File("newer.txt", new byte[] { 2 }, newer),
            ArchiveWriteEntry.File("unknown.txt", new byte[] { 3 })
        }, new ZipArchiveWriterOptions { Timestamp = epoch, ClampToTimestamp = true });

        using var runtime = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        Assert.Equal(older.DateTime, runtime.GetEntry("older.txt")!.LastWriteTime.DateTime);
        Assert.Equal(epoch.UtcDateTime, runtime.GetEntry("newer.txt")!.LastWriteTime.DateTime);
        Assert.Equal(epoch.UtcDateTime, runtime.GetEntry("unknown.txt")!.LastWriteTime.DateTime);
        Assert.Null(ArchiveTimestamp.ParseSourceDateEpoch(""));
        Assert.Throws<FormatException>(() => ArchiveTimestamp.ParseSourceDateEpoch("-1"));
        Assert.Throws<FormatException>(() => ArchiveTimestamp.ParseSourceDateEpoch("17e8"));
    }

    [Theory]
    [InlineData("../escape.txt")]
    [InlineData("/absolute.txt")]
    [InlineData("C:/drive.txt")]
    [InlineData("a/../../b.txt")]
    [InlineData("a//b.txt")]
    [InlineData("CON.txt")]
    [InlineData("name.")]
    public void UnsafeMemberNamesAreRejectedBeforeAnythingIsWritten(string name) {
        using var output = new MemoryStream();
        using var writer = new ZipArchiveWriter(output, leaveOpen: true);
        Assert.Throws<InvalidDataException>(() => writer.AddFile(name, new byte[] { 1 }));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void CollidingNamesAreRejected() {
        using var writer = new ZipArchiveWriter(new MemoryStream());
        writer.AddFile("Docs/a.txt", new byte[] { 1 });
        Assert.Throws<InvalidDataException>(() => writer.AddFile("docs/A.txt", new byte[] { 2 }));
        Assert.Throws<InvalidDataException>(() => writer.AddFile("Docs/a.txt/child", new byte[] { 2 }));
    }

    [Fact]
    public void DataDescriptorModeStreamsToNonseekableDestinations() {
        var output = new NonSeekableWriteStream();
        using (var writer = new ZipArchiveWriter(output, new ZipArchiveWriterOptions { UseDataDescriptors = true }, leaveOpen: true)) {
            using (Stream entry = writer.CreateEntry("docs/readme.txt")) entry.Write(Text, 0, Text.Length);
            writer.AddFile("stored.bin", new byte[] { 1, 2, 3 }, compression: ZipCompressionMethod.Stored);
        }
        byte[] zip = output.ToArray();
        Assert.Equal(0x08, zip[6] & 0x08);
        Assert.Equal(Text, ArchiveReader.ReadZip(zip).Single(file => file.RelativePath == "docs/readme.txt").Data);
        using var runtime = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        Assert.Equal(3, runtime.GetEntry("stored.bin")!.Length);
    }

    [Fact]
    public void ArchivesWithMoreThan65534MembersUseZip64EndRecords() {
        using var output = new MemoryStream();
        using (var writer = new ZipArchiveWriter(output, leaveOpen: true))
            for (int index = 0; index < 65_600; index++) writer.AddFile("f/" + index.ToString("D5"), Array.Empty<byte>());
        byte[] zip = output.ToArray();
        int end = zip.Length - 22;
        Assert.Equal(0xFFFF, BitConverter.ToUInt16(zip, end + 10));
        Assert.Equal(0x07064b50u, BitConverter.ToUInt32(zip, end - 20));
        using var runtime = new ZipArchive(new MemoryStream(zip), ZipArchiveMode.Read);
        Assert.Equal(65_600, runtime.Entries.Count);
        using var inspection = ArchiveInspector.Open(new MemoryStream(zip), new ArchiveInspectionOptions { MaxEntryCount = 70_000 });
        Assert.Equal(65_600, inspection.Entries.Count);
    }

    [Fact]
    public void FailedSourceLeavesNoCentralDirectory() {
        using var output = new MemoryStream();
        var writer = new ZipArchiveWriter(output, leaveOpen: true);
        writer.AddFile("good.txt", new byte[] { 1 });
        Assert.Throws<IOException>(() => writer.AddFile("bad.txt", new FailingStream()));
        Assert.Throws<InvalidOperationException>(() => writer.AddFile("later.txt", new byte[] { 2 }));
        writer.Dispose();
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadZip(output.ToArray()));
    }

    [Fact]
    public void DirectoryArchivesIgnoreFileTimesAndEnumerationOrderByDefault() {
        using var fixture = new ArchiveFixtures();
        string source = Path.Combine(fixture.Root, "source");
        Directory.CreateDirectory(Path.Combine(source, "b"));
        Directory.CreateDirectory(Path.Combine(source, "a", "empty"));
        File.WriteAllBytes(Path.Combine(source, "b", "two.txt"), Text);
        File.WriteAllBytes(Path.Combine(source, "a", "one.txt"), new byte[] { 1 });
        string first = Path.Combine(fixture.Root, "first.zip");
        ZipArchiveWriter.CreateFromDirectory(source, first);
        File.SetLastWriteTimeUtc(Path.Combine(source, "a", "one.txt"), new DateTime(2020, 5, 6, 7, 8, 10, DateTimeKind.Utc));
        string second = Path.Combine(fixture.Root, "second.zip");
        ZipArchiveWriter.CreateFromDirectory(source, second);

        Assert.Equal(File.ReadAllBytes(first), File.ReadAllBytes(second));
        Assert.Throws<IOException>(() => ZipArchiveWriter.CreateFromDirectory(source, second));
        using var runtime = new ZipArchive(File.OpenRead(first), ZipArchiveMode.Read);
        Assert.Equal(new[] { "a/", "a/empty/", "a/one.txt", "b/", "b/two.txt" }, runtime.Entries.Select(entry => entry.FullName));

        string preserved = Path.Combine(fixture.Root, "preserved.zip");
        ZipArchiveWriter.CreateFromDirectory(source, preserved, new ZipArchiveWriterOptions { PreserveFileTimes = true });
        using var withTimes = new ZipArchive(File.OpenRead(preserved), ZipArchiveMode.Read);
        Assert.Equal(new DateTime(2020, 5, 6, 7, 8, 10), withTimes.GetEntry("a/one.txt")!.LastWriteTime.DateTime);
    }

    internal static byte[] Write(IEnumerable<ArchiveWriteEntry> entries, ZipArchiveWriterOptions? options = null) {
        using var output = new MemoryStream();
        ZipArchiveWriter.Create(output, entries, options);
        return output.ToArray();
    }

    internal static string Hash(byte[] bytes) {
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(bytes).Take(8).Select(value => value.ToString("x2")));
    }

    internal sealed class NonSeekableStream : MemoryStream {
        internal NonSeekableStream(byte[] bytes) : base(bytes, writable: false) { }
        public override bool CanSeek => false;
    }

    internal sealed class NonSeekableWriteStream : MemoryStream {
        public override bool CanSeek => false;
    }

    internal sealed class FailingStream : Stream {
        private int _reads;
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) {
            if (_reads++ > 0) throw new IOException("Source failed.");
            buffer[offset] = 7;
            return 1;
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
