using System.IO.Compression;

namespace OfficeIMO.Archives.Tests;

public sealed class TarArchiveTests {
    private static readonly string LongPath = "docs/" + new string('d', 120) + "/" + new string('f', 90) + ".txt";
    private static readonly string VeryLongName = "docs/" + new string('n', 160) + ".txt";

    private static ArchiveWriteEntry[] Entries() => ZipArchiveWriterTests.SampleEntries().Concat(new[] {
        ArchiveWriteEntry.File(LongPath, Encoding.UTF8.GetBytes("split into the ustar prefix")),
        ArchiveWriteEntry.File(VeryLongName, Encoding.UTF8.GetBytes("needs a pax path"))
    }).ToArray();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrittenArchivesRoundTripThroughListingReadingAndExtraction(bool gzip) {
        byte[] tar = Write(Entries(), new TarArchiveWriterOptions { GZip = gzip });
        if (gzip) Assert.Equal(new byte[] { 0x1f, 0x8b }, tar.Take(2));
        else Assert.Equal(0, Write(Entries()).Length % 10240);

        var listing = ArchiveReader.ListTar(new MemoryStream(tar));
        Assert.Equal(new[] { "docs", "docs/readme.txt", "docs/zażółć-gęślą-jaźń.txt", "empty.bin", "data/random.bin", LongPath, VeryLongName },
            listing.Select(entry => entry.Name));
        Assert.Equal(ArchiveEntryKind.Directory, listing[0].Kind);
        Assert.Equal(0x1a4, listing[1].Mode);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(0), listing[1].LastWriteTime);

        var files = ArchiveReader.ReadTar(tar);
        Assert.Equal(ZipArchiveWriterTests.Text, files.Single(file => file.RelativePath == "docs/readme.txt").Data);
        Assert.Equal("needs a pax path", Encoding.UTF8.GetString(files.Single(file => file.RelativePath == VeryLongName).Data));

        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(tar), fixture.Destination);
        Assert.Equal(gzip ? ArchiveFormat.TarGZip : ArchiveFormat.Tar, result.Format);
        Assert.Equal(6, result.Files.Count);
        Assert.Equal("Zażółć gęślą jaźń.\n", File.ReadAllText(Path.Combine(fixture.Destination, "docs", "zażółć-gęślą-jaźń.txt")));
    }

    [Fact]
    public void ShortAsciiNamesUseUstarHeadersOnlyAndLongOrUnicodeNamesAddPaxRecords() {
        byte[] plain = Write(new[] { ArchiveWriteEntry.File("a.txt", new byte[] { 1 }) });
        Assert.Equal((byte)'0', plain[156]);
        Assert.Equal("ustar\000", Encoding.ASCII.GetString(plain, 257, 8));

        byte[] unicode = Write(new[] { ArchiveWriteEntry.File("zażółć.txt", new byte[] { 1 }) });
        Assert.Equal((byte)'x', unicode[156]);
        Assert.Contains("path=zażółć.txt\n", Encoding.UTF8.GetString(unicode, 512, 512));
    }

    [Fact]
    public void DefaultOutputIsByteIdenticalAcrossRunsAndTargets() {
        byte[] tar = Write(Entries());
        byte[] tgz = Write(Entries(), new TarArchiveWriterOptions { GZip = true });
        Assert.Equal(tar, Write(Entries()));
        Assert.Equal(tgz, Write(Entries(), new TarArchiveWriterOptions { GZip = true }));
        // The same digests are expected from every target framework and operating system.
        Assert.Equal("321a11e0cf250d4a", ZipArchiveWriterTests.Hash(tar));
        Assert.Equal("aec0fe300b2bcd8e", ZipArchiveWriterTests.Hash(tgz));
    }

    [Fact]
    public void LargeSizesAndOutOfRangeTimesUsePaxAndBase256() {
        var header = TarHeader("big.bin", '0', 0);
        Array.Clear(header, 124, 12); header[124] = 0x80; header[135] = 3;
        Assert.Equal(3, TarArchiveContent.ParseNumber(header, 124, 12, "size"));
        var negative = TarHeader("old.bin", '0', 0);
        for (int index = 136; index < 148; index++) negative[index] = 0xff;
        Assert.Equal(-1, TarArchiveContent.ParseNumber(negative, 136, 12, "mtime"));

        byte[] tar = Write(new[] { ArchiveWriteEntry.File("old.txt", new byte[] { 1 }, new DateTimeOffset(1960, 1, 1, 0, 0, 0, TimeSpan.Zero)) });
        Assert.Contains("mtime=-315619200\n", Encoding.UTF8.GetString(tar, 512, 512));
        Assert.Equal(new DateTimeOffset(1960, 1, 1, 0, 0, 0, TimeSpan.Zero), ArchiveReader.ListTar(new MemoryStream(tar)).Single().LastWriteTime);
    }

    [Fact]
    public void LinksAreListedAsMetadataAndRejectedOnlyWhenSelectedForExtraction() {
        byte[] tar = Build(
            (TarHeader("docs/target.txt", '0', 5), Encoding.ASCII.GetBytes("hello")),
            (TarHeader("docs/link.txt", '2', 0, link: "../../etc/passwd"), null),
            (TarHeader("docs/hard.txt", '1', 0, link: "docs/target.txt"), null),
            (TarHeader("docs/fifo", '6', 0), null));

        var listing = ArchiveReader.ListTar(new MemoryStream(tar));
        Assert.Equal(new[] { ArchiveEntryKind.File, ArchiveEntryKind.SymbolicLink, ArchiveEntryKind.HardLink, ArchiveEntryKind.Special },
            listing.Select(entry => entry.Kind));
        Assert.Equal("../../etc/passwd", listing[1].LinkTarget);
        Assert.Equal("docs/target.txt", Assert.Single(ArchiveReader.ReadTar(tar)).RelativePath);

        using var fixture = new ArchiveFixtures();
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(tar), fixture.Destination));
        Assert.Contains("symbolic link", error.Message);
        Assert.False(Directory.Exists(fixture.Destination));
        var result = ArchiveExtractor.Extract(fixture.Package(tar), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = path => path == "docs/target.txt" });
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "docs", "target.txt")));
        Assert.Single(result.Files);
    }

    [Theory]
    [InlineData("../evil.txt")]
    [InlineData("/etc/passwd")]
    [InlineData("docs/../../evil.txt")]
    [InlineData("C:/evil.txt")]
    public void UnsafeNamesAreRejectedOnReadAndExtraction(string name) {
        byte[] tar = Build((TarHeader(name, '0', 1), new byte[] { 1 }));
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(tar));
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(tar), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Equal(name, ArchiveReader.ListTar(new MemoryStream(tar)).Single().Name);
    }

    [Fact]
    public void PaxPathsAreValidatedLikeHeaderNames() {
        byte[] pax = Encoding.UTF8.GetBytes("22 path=../../outside\n");
        byte[] tar = Build((TarHeader("PaxHeaders/x", 'x', pax.Length), pax), (TarHeader("safe.txt", '0', 1), new byte[] { 1 }));
        Assert.Equal("../../outside", ArchiveReader.ListTar(new MemoryStream(tar)).Single().Name);
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(tar));
    }

    [Fact]
    public void CorruptAndTruncatedArchivesFailExplicitly() {
        byte[] valid = Build((TarHeader("a.txt", '0', 3), new byte[] { 1, 2, 3 }));
        byte[] badChecksum = (byte[])valid.Clone(); badChecksum[0] = (byte)'b';
        byte[] truncatedHeader = valid.Take(300).ToArray();
        byte[] truncatedData = valid.Take(514).ToArray();
        byte[] hugeDeclared = Build((TarHeader("a.txt", '0', 0), null));
        Octal(hugeDeclared, 124, 12, 07777777777L); Checksum(hugeDeclared);

        // A checksum mismatch or partial first block no longer looks like TAR at all.
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadTar(badChecksum));
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadTar(truncatedHeader));
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(truncatedData));
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(hugeDeclared, new ArchiveReadOptions { MaxEntryBytes = long.MaxValue }));
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadTar(valid, new ArchiveReadOptions { MaxEntryBytes = 2 }));

        byte[] secondBad = Build((TarHeader("a.txt", '0', 3), new byte[] { 1, 2, 3 }), (TarHeader("b.txt", '0', 0), null));
        secondBad[1024] = (byte)'c';
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(secondBad));
        byte[] malformedPax = Encoding.ASCII.GetBytes("99 path=x\n");
        Assert.Throws<InvalidDataException>(() => ArchiveReader.ReadTar(Build((TarHeader("p", 'x', malformedPax.Length), malformedPax),
            (TarHeader("a.txt", '0', 0), null))));
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadTar(Build((TarHeader("sparse", 'S', 0), null))));
    }

    [Fact]
    public void EntryCountPathDepthAndExpansionLimitsApply() {
        byte[] many = Build(Enumerable.Range(0, 5).Select(index => (TarHeader("f" + index, '0', 0), (byte[]?)null)).ToArray());
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadTar(many, new ArchiveReadOptions { MaxEntryCount = 4 }));
        byte[] deep = Build((TarHeader("a/b/c/d/e.txt", '0', 0), null));
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadTar(deep, new ArchiveReadOptions { MaxPathDepth = 4 }));
        byte[] two = Build((TarHeader("a", '0', 600), new byte[600]), (TarHeader("b", '0', 600), new byte[600]));
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadTar(two, new ArchiveReadOptions { MaxExpandedBytes = 1000 }));

        // A small .tar.gz cannot expand into an unbounded TAR stream.
        using var bomb = new MemoryStream();
        using (var writer = GZipArchive.CreateWriter(bomb, leaveOpen: true)) writer.Write(new byte[8 * 1024 * 1024], 0, 8 * 1024 * 1024);
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadTar(bomb.ToArray(),
            new ArchiveReadOptions { MaxExpandedBytes = 1024 * 1024, MaxEntryCount = 10 }));
    }

    [Fact]
    public void GnuLongNamesAndEmptyArchivesAreRead() {
        string name = "dir/" + new string('x', 150);
        byte[] longName = Encoding.UTF8.GetBytes(name + "\0");
        byte[] tar = Build((TarHeader("././@LongLink", 'L', longName.Length, gnu: true), longName),
            (TarHeader(name.Substring(0, 100), '0', 2, gnu: true), new byte[] { 4, 5 }));
        Assert.Equal(name, Assert.Single(ArchiveReader.ReadTar(tar)).RelativePath);
        Assert.Empty(ArchiveReader.ReadTar(new byte[10240]));
        Assert.Empty(ArchiveReader.ReadTar(Write(Array.Empty<ArchiveWriteEntry>(), new TarArchiveWriterOptions { GZip = true })));
    }

    [Fact]
    public void GZipMembersAreVerifiedAndConcatenated() {
        byte[] first = GZip(Encoding.ASCII.GetBytes("first "), new GZipWriterOptions { FileName = "report.txt", ModificationTime = DateTimeOffset.FromUnixTimeSeconds(1700000000) });
        byte[] joined = first.Concat(GZip(Encoding.ASCII.GetBytes("second"))).Concat(new byte[16]).ToArray();
        using var output = new MemoryStream();
        GZipReadResult result = GZipArchive.Decompress(new MemoryStream(joined), output);
        Assert.Equal("first second", Encoding.ASCII.GetString(output.ToArray()));
        Assert.Equal(2, result.MemberCount);
        Assert.Equal("report.txt", result.FileName);
        Assert.Equal(DateTimeOffset.FromUnixTimeSeconds(1700000000), result.ModificationTime);
        using (var runtime = new GZipStream(new MemoryStream(first), CompressionMode.Decompress)) {
            using var decoded = new MemoryStream(); runtime.CopyTo(decoded);
            Assert.Equal("first ", Encoding.ASCII.GetString(decoded.ToArray()));
        }

        byte[] badCrc = (byte[])first.Clone(); badCrc[badCrc.Length - 8] ^= 1;
        byte[] badSize = (byte[])first.Clone(); badSize[badSize.Length - 1] ^= 1;
        byte[] trailing = first.Concat(new byte[] { 1 }).ToArray();
        byte[] truncated = first.Take(first.Length - 3).ToArray();
        byte[] reserved = (byte[])first.Clone(); reserved[3] |= 0x20;
        foreach (byte[] bad in new[] { badCrc, badSize, trailing, truncated, reserved })
            Assert.Throws<InvalidDataException>(() => GZipArchive.Decompress(new MemoryStream(bad), Stream.Null));
        Assert.Throws<OfficeIMO.OfficeInspectionResourceLimitException>(() => GZipArchive.Decompress(new MemoryStream(joined), Stream.Null,
            new ArchiveReadOptions { MaxEntryBytes = 8 }));
    }

    [Fact]
    public void SingleFileGZipExtractsUnderItsStoredOrPackageName() {
        using var fixture = new ArchiveFixtures();
        string named = Path.Combine(fixture.Root, "named.gz");
        File.WriteAllBytes(named, GZip(new byte[] { 1, 2 }, new GZipWriterOptions { FileName = "inner.bin" }));
        var result = ArchiveExtractor.Extract(named, fixture.Destination);
        Assert.Equal(ArchiveFormat.GZip, result.Format);
        Assert.Equal("inner.bin", Assert.Single(result.Files).RelativePath);

        string unnamed = Path.Combine(fixture.Root, "notes.txt.gz");
        File.WriteAllBytes(unnamed, GZip(new byte[] { 3 }));
        Assert.Equal("notes.txt", Assert.Single(ArchiveExtractor.Extract(unnamed, Path.Combine(fixture.Root, "second")).Files).RelativePath);

        string hostile = Path.Combine(fixture.Root, "hostile.gz");
        File.WriteAllBytes(hostile, GZip(new byte[] { 4 }, new GZipWriterOptions { FileName = "../evil" }));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(hostile, Path.Combine(fixture.Root, "third")));
        Assert.Throws<NotSupportedException>(() => ArchiveInspector.Open(hostile));
    }

    [Theory]
    [InlineData("bsdtar-ustar.tar", ArchiveFormat.Tar)]
    [InlineData("bsdtar-pax.tar.gz", ArchiveFormat.TarGZip)]
    [InlineData("python-gnu.tar", ArchiveFormat.Tar)]
    public void IndependentProducersAreReadAndExtracted(string fixture, ArchiveFormat format) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentTar", fixture);
        var listing = ArchiveReader.ListTar(File.OpenRead(path));
        var link = listing.Single(entry => entry.Name == "docs/hello-link.txt");
        Assert.Equal(ArchiveEntryKind.SymbolicLink, link.Kind);
        Assert.Equal("hello.txt", link.LinkTarget);
        Assert.Equal(new DateTime(2024, 1, 2), listing.Single(entry => entry.Name == "docs/hello.txt").LastWriteTime!.Value.Date);

        var files = ArchiveReader.ReadTar(File.ReadAllBytes(path));
        Assert.Equal("Hello from an independent TAR producer.\n", Encoding.UTF8.GetString(files.Single(file => file.RelativePath == "docs/hello.txt").Data));
        if (fixture != "bsdtar-ustar.tar") {
            string longPath = "docs/a-very-long-directory-name-that-exceeds-the-classic-one-hundred-byte-ustar-name-field/and-a-nested-file-name.txt";
            Assert.Equal("Long path payload.\n", Encoding.UTF8.GetString(files.Single(file => file.RelativePath == longPath).Data));
        }
        if (fixture == "bsdtar-pax.tar.gz") {
            var polish = files.Single(file => file.RelativePath.Normalize(NormalizationForm.FormC) == "docs/zażółć-gęślą-jaźń.txt");
            Assert.Equal("Polish filename and payload: Zażółć gęślą jaźń.\n", Encoding.UTF8.GetString(polish.Data));
        }
        if (fixture == "python-gnu.tar") Assert.Equal("docs/hello.txt", listing.Single(entry => entry.Kind == ArchiveEntryKind.HardLink).LinkTarget);

        using var output = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(path, output.Destination,
            new ArchiveExtractionOptions { IncludeFile = name => !name.Contains("link") });
        Assert.Equal(format, result.Format);
        Assert.Equal(files.Count, result.Files.Count);
    }

    [Fact]
    public void IndependentMultiMemberGZipIsConcatenated() {
        using var output = new MemoryStream();
        var result = GZipArchive.Decompress(File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentTar", "gzip-multimember.gz")), output);
        Assert.Equal("first member\nsecond member\n", Encoding.ASCII.GetString(output.ToArray()));
        Assert.Equal(2, result.MemberCount);
    }

    internal static byte[] Write(IEnumerable<ArchiveWriteEntry> entries, TarArchiveWriterOptions? options = null) {
        using var output = new MemoryStream();
        TarArchiveWriter.Create(output, entries, options);
        return output.ToArray();
    }

    private static byte[] GZip(byte[] content, GZipWriterOptions? options = null) {
        using var output = new MemoryStream();
        GZipArchive.Compress(new MemoryStream(content), output, options);
        return output.ToArray();
    }

    private static byte[] TarHeader(string name, char type, long size, string? link = null, bool gnu = false) {
        var block = new byte[512];
        byte[] nameBytes = Encoding.UTF8.GetBytes(name);
        Array.Copy(nameBytes, block, Math.Min(100, nameBytes.Length));
        Octal(block, 100, 8, 0x1a4); Octal(block, 108, 8, 0); Octal(block, 116, 8, 0);
        Octal(block, 124, 12, size); Octal(block, 136, 12, 1700000000);
        block[156] = (byte)type;
        if (link != null) Encoding.UTF8.GetBytes(link).CopyTo(block, 157);
        Encoding.ASCII.GetBytes(gnu ? "ustar  \0" : "ustar\000").CopyTo(block, 257);
        Checksum(block);
        return block;
    }

    private static byte[] Build(params (byte[] Header, byte[]? Data)[] members) {
        var output = new MemoryStream();
        foreach (var (header, data) in members) {
            output.Write(header, 0, header.Length);
            if (data == null) continue;
            output.Write(data, 0, data.Length);
            output.Write(new byte[(512 - data.Length % 512) % 512], 0, (512 - data.Length % 512) % 512);
        }
        output.Write(new byte[1024], 0, 1024);
        return output.ToArray();
    }

    private static void Octal(byte[] block, int offset, int length, long value) {
        string text = Convert.ToString(value, 8).PadLeft(length - 1, '0');
        Encoding.ASCII.GetBytes(text).CopyTo(block, offset);
        block[offset + length - 1] = 0;
    }

    private static void Checksum(byte[] block) {
        for (int index = 148; index < 156; index++) block[index] = (byte)' ';
        Octal(block, 148, 7, block.Take(512).Sum(value => (long)value));
    }
}
