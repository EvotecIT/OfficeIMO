namespace OfficeIMO.Archives.Tests;

public sealed class TarWrapperTests {
    [Theory]
    [InlineData("archive.tar.zst", ArchiveFormat.TarZstandard)]
    [InlineData("archive.tar.xz", ArchiveFormat.TarXz)]
    [InlineData("archive.tar.lz", ArchiveFormat.TarLzip)]
    public void CompressedTarsReadListAndExtractLikePlainTar(string fixture, ArchiveFormat format) {
        IReadOnlyList<ArchiveEntryData> expected = ArchiveReader.ReadTar(ReadFixture("archive.tar"));
        IReadOnlyList<ArchiveEntryData> files = ArchiveReader.ReadTar(ReadFixture(fixture));
        Assert.Equal(expected.Select(file => file.RelativePath), files.Select(file => file.RelativePath));
        for (int index = 0; index < files.Count; index++) Assert.Equal(expected[index].Data, files[index].Data);

        TarArchiveEntry link = ArchiveReader.ListTar(new MemoryStream(ReadFixture(fixture))).Single(entry => entry.Kind == ArchiveEntryKind.SymbolicLink);
        Assert.Equal("text.txt", link.LinkTarget);

        using var output = new ArchiveFixtures();
        string package = Path.Combine(output.Root, fixture);
        File.WriteAllBytes(package, ReadFixture(fixture));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, output.Destination,
            new ArchiveExtractionOptions { IncludeFile = name => !name.Contains("link") });
        Assert.Equal(format, result.Format);
        Assert.Equal(expected.Select(file => file.RelativePath), result.Files.Select(file => file.RelativePath).OrderBy(path => path, StringComparer.Ordinal));
        Assert.Equal(expected.Single(file => file.RelativePath == "docs/mixed.bin").Data, File.ReadAllBytes(Path.Combine(output.Destination, "docs", "mixed.bin")));
    }

    [Fact]
    public void DeclaredWindowsAboveTheLimitAreRefused() {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ArchiveReader.ReadTar(ReadFixture("archive-window128m.tar.zst")));
        Assert.Equal(nameof(ArchiveReadOptions.MaxWindowBytes), error.LimitName);

        IReadOnlyList<ArchiveEntryData> files = ArchiveReader.ReadTar(ReadFixture("archive-window128m.tar.zst"),
            new ArchiveReadOptions { MaxWindowBytes = 128L * 1024 * 1024 });
        Assert.Equal(2, files.Count);

        using var output = new ArchiveFixtures();
        string package = Path.Combine(output.Root, "archive.tar.zst");
        File.WriteAllBytes(package, ReadFixture("archive-window128m.tar.zst"));
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveExtractor.Extract(package, output.Destination));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, output.Destination,
            new ArchiveExtractionOptions { MaxWindowBytes = 128L * 1024 * 1024, IncludeFile = name => !name.Contains("link") });
        Assert.Equal(ArchiveFormat.TarZstandard, result.Format);
    }

    [Theory]
    [InlineData("archive.tar.zst")]
    [InlineData("archive.tar.xz")]
    [InlineData("archive.tar.lz")]
    public void DecompressedTarIsBoundedByTheExpansionLimit(string fixture) {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(() =>
            ArchiveReader.ReadTar(ReadFixture(fixture), new ArchiveReadOptions { MaxExpandedBytes = 64 * 1024, MaxEntryCount = 4 }));
        Assert.Equal(nameof(ArchiveReadOptions.MaxExpandedBytes), error.LimitName);
    }

    [Theory]
    [InlineData("archive.tar.zst", "Zstandard")]
    [InlineData("archive.tar.xz", "xz")]
    [InlineData("archive.tar.lz", "lzip")]
    public void TruncatedCompressedTarsFailWithTheCodecError(string fixture, string codec) {
        byte[] whole = ReadFixture(fixture);
        byte[] truncated = whole.Take(whole.Length / 2).ToArray();
        ArchiveCorruptionException error = Assert.Throws<ArchiveCorruptionException>(() => ArchiveReader.ReadTar(truncated));
        Assert.Equal(codec, error.Format);
    }

    [Theory]
    [InlineData("IndependentZstandard", "text-level3-check.zst", ".zst", ArchiveFormat.Zstandard)]
    [InlineData("IndependentXzLzip", "text-preset6-crc64.xz", ".xz", ArchiveFormat.Xz)]
    [InlineData("IndependentXzLzip", "text-level6.lz", ".lz", ArchiveFormat.Lzip)]
    public void SingleCompressedFilesExtractUnderThePackageName(string folder, string fixture, string suffix, ArchiveFormat format) {
        string source = Path.Combine(AppContext.BaseDirectory, "Fixtures", folder);
        using var output = new ArchiveFixtures();
        string package = Path.Combine(output.Root, "harbour-log.txt" + suffix);
        File.Copy(Path.Combine(source, fixture), package);
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, output.Destination);
        Assert.Equal(format, result.Format);
        Assert.Equal("harbour-log.txt", Assert.Single(result.Files).RelativePath);
        Assert.Equal(File.ReadAllBytes(Path.Combine(source, "inputs", "text.txt")), File.ReadAllBytes(Path.Combine(output.Destination, "harbour-log.txt")));
        Assert.Throws<NotSupportedException>(() => ArchiveInspector.Open(package));
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadTar(File.ReadAllBytes(package)));
    }

    [Fact]
    public void AllowedFormatsAcceptTheNewFormats() {
        using var output = new ArchiveFixtures();
        string package = Path.Combine(output.Root, "archive.tar.xz");
        File.WriteAllBytes(package, ReadFixture("archive.tar.xz"));
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(package, output.Destination,
            new ArchiveExtractionOptions { AllowedFormats = new[] { ArchiveFormat.TarGZip } }));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, output.Destination,
            new ArchiveExtractionOptions { AllowedFormats = new[] { ArchiveFormat.TarXz, ArchiveFormat.Lzip }, IncludeFile = name => !name.Contains("link") });
        Assert.Equal(ArchiveFormat.TarXz, result.Format);
    }

    private static byte[] ReadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentTarWrappers", name));
}
