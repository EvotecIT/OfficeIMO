namespace OfficeIMO.Archives.Tests;

public sealed class MsiSequenceValidationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CabinetOrderMustMatchFileSequenceEvenForExcludedFiles(bool exclude) {
        using var fixture = new ArchiveFixtures();
        byte[] package = ArchiveFixtures.MsiFiles(false, true, 1, null,
            new MsiFixtureFile("First.adml"), new MsiFixtureFile("Second.adml", "world"));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(package), fixture.Destination, new ArchiveExtractionOptions { IncludeFile = _ => !exclude }));
        Assert.Contains("sequence", error.Message);
        AssertNoOutput(fixture);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstMediaDiskMustBeOneEvenForExcludedFiles(bool exclude) {
        using var fixture = new ArchiveFixtures();
        byte[] package = ArchiveFixtures.MsiFiles(false, false, 2, null, new MsiFixtureFile("Template.adml"));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(package), fixture.Destination, new ArchiveExtractionOptions { IncludeFile = _ => !exclude }));
        Assert.Contains("DiskId", error.Message);
        AssertNoOutput(fixture);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FileTableRowOrderDoesNotChangeAuthoredCabinetSequence(bool reverseRows) {
        using var fixture = new ArchiveFixtures();
        byte[] package = ArchiveFixtures.MsiFiles(reverseRows, false, 1, null,
            new MsiFixtureFile("First.adml"), new MsiFixtureFile("Second.adml", "world"),
            new MsiFixtureFile("Third.adml", "third", cabinet: 3));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination);
        Assert.Equal(3, result.Files.Count);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/First.adml")));
        Assert.Equal("world", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Second.adml")));
        Assert.Equal("third", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Third.adml")));
    }

    private static void AssertNoOutput(ArchiveFixtures fixture) {
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(true, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void MediaOnlyCabinetsObserveCountAndMemberValidation(bool exclude, int defect) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = defect == 2 ? new byte[36] : ArchiveFixtures.StoredCab("Unknown", Encoding.UTF8.GetBytes("extra"));
        byte[] package = ArchiveFixtures.Msi(unmappedCabinet: cabinet);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxCabinetCount = defect == 0 ? 1 : 2, IncludeFile = _ => !exclude }));
        if (defect == 0) Assert.Contains("embedded-cabinet count", error.Message);
        AssertNoOutput(fixture);
    }

    [Fact]
    public void ValidEmptyMediaCabinetIsAcceptedAndStillCountsTowardLimits() {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = ArchiveFixtures.StoredCab();
        Array.Resize(ref cabinet, 44);
        Buffer.BlockCopy(BitConverter.GetBytes(cabinet.Length), 0, cabinet, 8, 4);
        cabinet[40] = 0; cabinet[41] = 0;
        byte[] package = ArchiveFixtures.Msi(unmappedCabinet: cabinet);
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
            new ArchiveExtractionOptions { MaxCabinetCount = 2 });
        Assert.Single(result.Files);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions/en-US/Template.adml")));
    }
}
