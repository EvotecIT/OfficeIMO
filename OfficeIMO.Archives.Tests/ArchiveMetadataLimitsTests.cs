using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveMetadataLimitsTests {
    [Theory]
    [InlineData(4, false)]
    [InlineData(12, false)]
    [InlineData(20, false)]
    [InlineData(4, true)]
    [InlineData(12, true)]
    [InlineData(20, true)]
    public void RejectsNonzeroFixedCabinetReservedFieldsBeforeSelection(int offset, bool excluded) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = ArchiveFixtures.StoredCab("Template.admx", Encoding.UTF8.GetBytes("hello"));
        cabinet[offset + 3] = 1;

        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(cabinet), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => !excluded }));

        Assert.Contains("CAB_HEADER", error.Message);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void TableRowLimitDoesNotLimitTheStringPool() {
        using var fixture = new ArchiveFixtures();
        // _Columns has 23 rows; this ordinary one-file package has more than 23 strings.
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi()),
            fixture.Destination, new ArchiveExtractionOptions { MaxTableRows = 23 });

        Assert.Single(result.Files);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StringPoolHasItsOwnInclusiveEntryLimit(bool exceedsLimit) {
        using var fixture = new ArchiveFixtures();
        byte[] package = ArchiveFixtures.Msi();
        Assert.True(OfficeCompoundFileReader.TryRead(package, out Dictionary<string, byte[]> streams, out string? error), error);
        // This fixture uses ordinary four-byte records; the first record is the pool header.
        int stringCount = streams[MsiPackageReader.EncodeStreamName("_StringPool", true)].Length / 4 - 1;
        Assert.True(stringCount > 23);
        var options = new ArchiveExtractionOptions {
            MaxTableRows = 23, MaxStringPoolEntries = stringCount - (exceedsLimit ? 1 : 0)
        };

        if (exceedsLimit) {
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination, options));
            Assert.Contains("string-pool entry limit", exception.Message);
            Assert.False(Directory.Exists(fixture.Destination));
        } else {
            Assert.Single(ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination, options).Files);
        }
    }

    [Fact]
    public void TableRowLimitStillRejectsAnOversizedTable() {
        using var fixture = new ArchiveFixtures();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi()), fixture.Destination, new ArchiveExtractionOptions { MaxTableRows = 22 }));
        Assert.Contains("row limit: _Columns", error.Message);
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void StringPoolLimitMustBePositive(int limit) {
        using var fixture = new ArchiveFixtures();
        ArgumentOutOfRangeException error = Assert.Throws<ArgumentOutOfRangeException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi()), fixture.Destination, new ArchiveExtractionOptions { MaxStringPoolEntries = limit }));
        Assert.Equal(nameof(ArchiveExtractionOptions.MaxStringPoolEntries), error.ParamName);
    }
}
