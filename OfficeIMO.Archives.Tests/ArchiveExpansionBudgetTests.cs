using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveExpansionBudgetTests {
    [Theory]
    [InlineData(true, 0)]
    [InlineData(false, 0)]
    [InlineData(true, 1)]
    [InlineData(false, 1)]
    public void ExactExpansionBudgetAllowsFollowingEmptyCabinet(bool selected, int emptyCompression) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.Msi(twoCabinets: true, emptySecondPayload: true, emptySecondCompression: emptyCompression));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 5, IncludeFile = _ => selected });
        Assert.Equal(selected ? 5 : 0, result.TotalBytes);
        Assert.Equal(selected ? 2 : 0, result.Files.Count);
        if (selected) {
            Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
            Assert.Equal(0, new FileInfo(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Second.adml")).Length);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExhaustedBudgetRejectsFollowingNonemptyCabinetEvenWhenUnselected(bool selected) {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)), fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 5, IncludeFile = _ => selected }));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(0x0f03)]
    public void ZeroRemainingExpansionAllowsSupportedFoldersWithNoDataRecords(int compression) {
        IReadOnlyList<OfficeCabinetEntry> entries = OfficeCabinetArchiveReader.Read(ArchiveFixtures.EmptyStoredCab("empty", compression),
            0, 1, 1, out long expandedBytes);
        Assert.Equal(0, expandedBytes);
        Assert.Empty(Assert.Single(entries).Data);
    }

    [Fact]
    public void ExpandedLimitsStillRejectNegativeInternalAndZeroPublicBudgets() {
        byte[] cabinet = ArchiveFixtures.EmptyStoredCab("empty");
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeCabinetArchiveReader.Read(cabinet, -1, 1, 1));
        using var fixture = new ArchiveFixtures();
        Assert.Throws<ArgumentOutOfRangeException>(() => ArchiveExtractor.Extract(fixture.Package(cabinet), fixture.Destination,
            new ArchiveExtractionOptions { MaxExpandedBytes = 0 }));
        Assert.False(Directory.Exists(fixture.Destination));
    }
}
