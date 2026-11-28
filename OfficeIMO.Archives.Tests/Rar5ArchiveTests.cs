namespace OfficeIMO.Archives.Tests;

public sealed partial class Rar5ArchiveTests {
    [Fact]
    public void ReadsStoredMembersFromTheAuthoritativeFileHeadersAndRefusesLinks() {
        byte[] source = Fixture("stored.rar");
        using var stream = new MemoryStream(source, writable: false);
        var entries = ArchiveReader.ListRar5(stream);

        Rar5ArchiveEntry target = Assert.Single(entries, entry => entry.RelativePath == "notes/café.txt");
        Assert.Equal(0, target.CompressionMethod);
        Assert.Equal(1280, target.Length);
        Assert.True(target.CanRead);
        Assert.Equal(ArchiveEntryKind.Directory,
            Assert.Single(entries, entry => entry.RelativePath == "notes/empty-dir").Kind);
        Assert.Equal(ArchiveEntryKind.SymbolicLink,
            Assert.Single(entries, entry => entry.RelativePath == "notes/shortcut").Kind);
        Assert.Equal(ArchiveEntryKind.File,
            Assert.Single(entries, entry => entry.RelativePath == "notes/empty.txt").Kind);
        Assert.Throws<NotSupportedException>(() => ArchiveReader.ReadRar5(source));

        var budget = new ArchiveOperationBudget();
        var files = ArchiveReader.ReadRar5(source, new ArchiveReadOptions {
            IncludeFile = path => path == target.RelativePath, OperationBudget = budget, MaxExpandedBytes = target.Length
        });
        ArchiveEntryData selected = Assert.Single(files);
        Assert.Equal(target.RelativePath, selected.RelativePath);
        Assert.Equal(string.Concat(Enumerable.Repeat("The north quay.\n", 80)), Encoding.UTF8.GetString(selected.Data));
        Assert.Equal(selected.Length, budget.Usage.SelectedBytes);
        Assert.Equal(entries.Count, budget.Usage.EntryCount);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void ReportsTheAuthoredCompressionProfileForASelectedMember(int method) {
        byte[] data = Fixture($"compressed-{method}.rar");
        using var input = new MemoryStream(data, writable: false);
        Rar5ArchiveEntry entry = Assert.Single(ArchiveReader.ListRar5(input));
        Assert.Equal(method, entry.CompressionMethod);
        Assert.Equal(128 * 1024, entry.DictionarySize);
        Assert.False(entry.CanRead);
        var error = Assert.Throws<Rar5CompressionNotSupportedException>(() => ArchiveReader.ReadRar5(data));
        Assert.Equal(entry.Name, error.EntryName);
        Assert.Equal(method, error.Method);
        Assert.Equal(entry.DictionarySize, error.DictionarySize);
        Assert.Contains(entry.Name, error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("encrypted-headers.rar", null)]
    [InlineData("encrypted-member.rar", "notes/café.txt")]
    public void IdentifiesHeaderAndMemberEncryption(string name, string? encryptedName) {
        using var input = new MemoryStream(Fixture(name), writable: false);
        var error = Assert.Throws<Rar5EncryptionNotSupportedException>(() => ArchiveReader.ListRar5(input));
        Assert.Equal(encryptedName, error.EntryName);
        Assert.True(error.ByteOffset >= 8);
    }

    [Theory]
    [InlineData("legacy.rar", 0)]
    [InlineData("legacy-prefixed.sfx", 31)]
    public void RejectsRar4AtItsSignatureOffset(string name, int offset) {
        using var input = new MemoryStream(Fixture(name), writable: false);
        var error = Assert.Throws<NotSupportedException>(() => ArchiveReader.ListRar5(input));
        Assert.Contains("RAR 4 not supported", error.Message, StringComparison.Ordinal);
        Assert.Contains($"byte offset {offset}", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractsAStoredFileAndRollsBackOnPayloadCorruption() {
        byte[] bytes = Fixture("stored.rar");
        var options = new ArchiveExtractionOptions {
            AllowedFormats = new[] { ArchiveFormat.Rar5 },
            IncludeFile = path => path == "notes/café.txt"
        };
        string expected = string.Concat(Enumerable.Repeat("The north quay.\n", 80));
        using var good = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(good.Package(bytes), good.Destination, options);
        Assert.Equal(ArchiveFormat.Rar5, result.Format);
        Assert.Equal(expected, File.ReadAllText(Path.Combine(good.Destination, Assert.Single(result.Files).RelativePath)));

        byte[] stored = Encoding.UTF8.GetBytes(expected);
        int start = Enumerable.Range(0, bytes.Length - stored.Length + 1)
            .Single(offset => bytes.Skip(offset).Take(stored.Length).SequenceEqual(stored));
        bytes[start] ^= 1;
        using var bad = new ArchiveFixtures();
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(bad.Package(bytes), bad.Destination, options));
        Assert.Contains("file checksum", error.Message, StringComparison.Ordinal);
        Assert.False(Directory.Exists(bad.Destination));
        Assert.Empty(Directory.GetDirectories(bad.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData("MaxInputBytes")]
    [InlineData("MaxEntryCount")]
    [InlineData("MaxPathDepth")]
    [InlineData("MaxExpandedBytes")]
    public void EnforcesBudgetsBeforeReturningASelectedMember(string limit) {
        byte[] data = Fixture("stored.rar");
        var options = new ArchiveReadOptions { IncludeFile = path => path == "notes/café.txt" };
        switch (limit) {
            case "MaxInputBytes": options.MaxInputBytes = data.Length - 1; break;
            case "MaxEntryCount": options.MaxEntryCount = 5; break;
            case "MaxPathDepth": options.MaxPathDepth = 1; break;
            case "MaxExpandedBytes": options.MaxExpandedBytes = 1279; break;
        }
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadRar5(data, options));
        Assert.Equal(limit, error.LimitName);
    }

    [Fact]
    public void HeaderListingObservesSharedWorkAllowanceWithoutReadingFileData() {
        using var input = new MemoryStream(Fixture("stored.rar"), writable: false);
        var options = new ArchiveReadOptions {
            OperationBudget = new ArchiveOperationBudget(new ArchiveOperationLimits { MaxDecompressionWorkBytes = 1024 })
        };
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveReader.ListRar5(input, options));
        Assert.Equal("Operation.MaxDecompressionWorkBytes", error.LimitName);
        Assert.Contains("byte offset 8", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PreCancelledStreamCallsDoNotCheckArchiveDepth() {
        using var input = new MemoryStream(Fixture("stored.rar"), writable: false);
        var options = new ArchiveReadOptions {
            OperationBudget = new ArchiveOperationBudget(new ArchiveOperationLimits { MaxDepth = 0 }).CreateNested()
        };
        var cancelled = new CancellationToken(canceled: true);
        Assert.ThrowsAny<OperationCanceledException>(() => ArchiveReader.ListRar5(input, options, cancelled));
        Assert.ThrowsAny<OperationCanceledException>(() => ArchiveReader.ReadRar5(input, options, cancelled));
        Assert.Equal(0, input.Position);
    }

    private static byte[] Fixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentRar5", name));
}
