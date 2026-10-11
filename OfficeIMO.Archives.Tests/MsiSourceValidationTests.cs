namespace OfficeIMO.Archives.Tests;

public sealed class MsiSourceValidationTests {
    [Theory]
    [InlineData(null, 0)]
    [InlineData(0, 0)]
    [InlineData(2, 8192)]
    public void LooseDefaultOrExplicitSourceIsUnsupportedEvenWhenFiltered(int? wordCount, int attributes) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.MsiAliasFiles(false, wordCount,
            new MsiFixtureFile("Template.adml", attributes: attributes)));
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [Theory]
    [InlineData(2, 0)]
    [InlineData(0, 16384)]
    public void CabinetSourcesHonorSummaryDefaultAndExplicitCompressedOverride(int wordCount, int attributes) {
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false, wordCount,
            new MsiFixtureFile("Template.adml", attributes: attributes))), fixture.Destination);
        Assert.Equal(5, Assert.Single(result.Files).Length);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "en-US", "Template.adml")));
    }

    [Fact]
    public void ContradictoryCompressionOverridesAreInvalid() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false, 2,
            new MsiFixtureFile("Template.adml", attributes: 0x6000))), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void FilteredSingletonCabinetLengthMustMatchFileSize() {
        using var fixture = new ArchiveFixtures();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(wrongSize: true)), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.Contains("FileSize", error.Message);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }
}
