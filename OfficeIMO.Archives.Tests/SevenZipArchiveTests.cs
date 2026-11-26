using System.Security.Cryptography;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Tests;

public sealed class SevenZipArchiveTests {
    [Theory]
    [InlineData("solid-lzma2.7z", ArchiveCompression.Lzma2, true)]
    [InlineData("independent-lzma.7z", ArchiveCompression.Lzma, false)]
    public void ReadsProducerArchivesWithTheirRealFolderAndMemberBoundaries(string fixture, ArchiveCompression compression, bool solid) {
        byte[] source = ReadFixture(fixture);
        using var stream = new MemoryStream(source, writable: false);
        var entries = ArchiveReader.ListSevenZip(stream);
        Assert.Equal(ArchiveEntryKind.Directory, Assert.Single(entries, entry => entry.RelativePath == "notes/empty-dir").Kind);
        SevenZipArchiveEntry empty = Assert.Single(entries, entry => entry.RelativePath == "notes/empty.txt");
        Assert.Equal(ArchiveEntryKind.File, empty.Kind);
        Assert.Equal(0, empty.Length);
        Assert.Null(empty.FolderOrdinal);

        SevenZipArchiveEntry first = Assert.Single(entries, entry => entry.RelativePath == "notes/café.txt");
        SevenZipArchiveEntry second = Assert.Single(entries, entry => entry.RelativePath == "notes/second.txt");
        Assert.Equal(compression, first.Compression);
        Assert.Equal(compression, second.Compression);
        Assert.Equal(solid, first.FolderOrdinal == second.FolderOrdinal);

        var budget = new ArchiveOperationBudget();
        var selected = ArchiveReader.ReadSevenZip(source, new ArchiveReadOptions {
            OperationBudget = budget,
            MaxExpandedBytes = solid ? 64L * 1024 * 1024 : first.Length,
            IncludeFile = path => path == "notes/café.txt"
        });
        ArchiveEntryData file = Assert.Single(selected);
        Assert.Equal("notes/café.txt", file.RelativePath);
        Assert.Equal(string.Concat(Enumerable.Repeat("Harbour lights, chapter two.\n", 800)), Encoding.UTF8.GetString(file.Data));
        Assert.Equal(file.Length, budget.Usage.SelectedBytes);
        if (solid) {
            Assert.True(budget.Usage.ExpandedBytes > 2 * 1024 * 1024);
            ArchiveEntryData multiChunk = Assert.Single(ArchiveReader.ReadSevenZip(source, new ArchiveReadOptions {
                IncludeFile = path => path == "multi.txt"
            }));
            Assert.Equal("6851f897b10d1d714125ce8f22402779a1655c476cd41813e34fbaab7a9eed2d", ContentHash(multiChunk.Data));
        }
    }

    [Fact]
    public void ContinuesAnLzma2StreamAcrossAStoredChunkWithoutResettingItsModels() {
        ArchiveEntryData member = Assert.Single(ArchiveReader.ReadSevenZip(ReadFixture("mixed-lzma2.7z")));

        Assert.Equal("mixed.bin", member.RelativePath);
        Assert.Equal("091835e37b6448d55fac13bdc704637ef616b52f80cb620af73b0ad2e6d0fd16", ContentHash(member.Data));
    }

    [Fact]
    public void ListsAndReadsAnArchiveWithOnlyAnEmptyFile() {
        byte[] bytes = ReadFixture("copy.7z");
        using var stream = new MemoryStream(bytes, writable: false);
        SevenZipArchiveEntry member = Assert.Single(ArchiveReader.ListSevenZip(stream));
        Assert.Equal(ArchiveEntryKind.File, member.Kind);
        Assert.Equal("notes/empty.txt", member.RelativePath);
        Assert.Empty(Assert.Single(ArchiveReader.ReadSevenZip(bytes)).Data);
    }

    [Fact]
    public void ExtractsStoredDataAndRefusesToPublishAFileWithAnInvalidChecksum() {
        byte[] bytes = ReadFixture("stored.7z");
        using var good = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(good.Package(bytes), good.Destination,
            new ArchiveExtractionOptions { AllowedFormats = new[] { ArchiveFormat.SevenZip } });
        Assert.Equal(ArchiveFormat.SevenZip, result.Format);
        string relative = Assert.Single(result.Files).RelativePath;
        Assert.Equal("notes/café.txt", relative);
        Assert.Equal(string.Concat(Enumerable.Repeat("Harbour lights, chapter two.\n", 800)),
            File.ReadAllText(Path.Combine(good.Destination, relative)));

        bytes[32] ^= 1;
        using var bad = new ArchiveFixtures();
        var error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(bad.Package(bytes), bad.Destination));
        Assert.Contains("checksum", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(Directory.Exists(bad.Destination));
        Assert.Empty(Directory.GetDirectories(bad.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData("encrypted-header.7z")]
    [InlineData("encrypted-member.7z")]
    public void RefusesEncryptionDuringInventory(string fixture) {
        using var stream = new MemoryStream(ReadFixture(fixture), writable: false);
        var error = Assert.Throws<NotSupportedException>(() => ArchiveReader.ListSevenZip(stream));
        Assert.Contains("Encrypted 7z", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("MaxInputBytes")]
    [InlineData("MaxEntryCount")]
    [InlineData("MaxPathDepth")]
    [InlineData("MaxExpandedBytes")]
    [InlineData("MaxDecompressionWorkBytes")]
    public void AppliesLimitsEvenWhenOnlyOneSmallMemberIsSelected(string limit) {
        byte[] data = ReadFixture("solid-lzma2.7z");
        var options = new ArchiveReadOptions { IncludeFile = path => path == "notes/café.txt" };
        switch (limit) {
            case "MaxInputBytes": options.MaxInputBytes = data.Length - 1; break;
            case "MaxEntryCount": options.MaxEntryCount = 5; break;
            case "MaxPathDepth": options.MaxPathDepth = 1; break;
            case "MaxExpandedBytes": options.MaxExpandedBytes = 1024 * 1024; break;
            case "MaxDecompressionWorkBytes": options.MaxDecompressionWorkBytes = 1024 * 1024; break;
        }
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveReader.ReadSevenZip(data, options));
        Assert.Equal(limit, error.LimitName);
    }

    [Fact]
    public void ListingRejectsCopyWithDifferentPackedAndUnpackedSizes() {
        byte[] bytes = ReadFixture("stored.7z");
        RewriteNextHeader(bytes, (data, start) => {
            // The producer's first packed length is a three-byte integer in this plain header.
            Assert.Equal(0xc0, data[start + 6]);
            data[start + 7]--;
        });
        using var stream = new MemoryStream(bytes, writable: false);

        var error = Assert.Throws<InvalidDataException>(() => ArchiveReader.ListSevenZip(stream));
        Assert.Contains("Copy coder", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractionBoundsTheRetainedPathIndexBeforePublishingFiles() {
        using var fixture = new ArchiveFixtures();
        var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveExtractor.Extract(
            fixture.Package(ReadFixture("copy.7z")), fixture.Destination,
            new ArchiveExtractionOptions { MaxMetadataBytes = 100 }));

        // This fixture's header fits in 100 bytes, but the retained directory and file entries exceed the allowance.
        Assert.Equal("MaxMetadataBytes", error.LimitName);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    private static string ContentHash(byte[] data) {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
    }

    private static void RewriteNextHeader(byte[] bytes, Action<byte[], int> mutate) {
        using var stream = new MemoryStream(bytes);
        using var reader = new BinaryReader(stream);
        stream.Position = 12;
        int start = 32 + checked((int)reader.ReadUInt64());
        int length = checked((int)reader.ReadUInt64());
        mutate(bytes, start);
        using var writer = new BinaryWriter(stream);
        writer.Write(OfficePngCrc32.Compute(bytes, start, length));
        stream.Position = 8;
        writer.Write(OfficePngCrc32.Compute(bytes, 12, 20));
    }

    private static byte[] ReadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentSevenZip", name));
}
