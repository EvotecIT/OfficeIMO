namespace OfficeIMO.Archives.Tests;

public sealed class MsiPayloadAliasTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void IdenticalAliasesPublishOneDeterministicFileAcrossTableAndCabinetOrder(bool reverseRows, bool separateCabinets) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.MsiAliasFiles(reverseRows,
            new MsiFixtureFile("fidoauth.adml"), new MsiFixtureFile("FidoAuth.adml", cabinet: separateCabinets ? 2 : 1),
            new MsiFixtureFile("FidoAuth.adml", cabinet: separateCabinets ? 3 : 1)));
        var selected = new List<string>();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = path => { selected.Add(path); return true; } });
        Assert.Equal("PolicyDefinitions/en-US/FidoAuth.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal(new[] { "PolicyDefinitions/en-US/FidoAuth.adml" }, selected);
        Assert.Equal(5, result.TotalBytes);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "en-US", "FidoAuth.adml")));
        Assert.Single(Directory.GetFiles(fixture.Destination, "*", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConflictingAliasBytesCannotPublishEvenWhenTheLogicalFileIsExcluded(bool exclude) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
            new MsiFixtureFile("Template.adml"), new MsiFixtureFile("Template.adml", cabinet: 2),
            new MsiFixtureFile("Template.adml", "world", cabinet: 3)));
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => !exclude }));
        Assert.Contains("payload", error.Message.ToLowerInvariant());
        AssertNoOutput(fixture);
    }

    [Fact]
    public void ExcludedIdenticalAliasesAreVerifiedWithoutPublishingOrInflatingResultBytes() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
            new MsiFixtureFile("Template.adml"), new MsiFixtureFile("Template.adml", cabinet: 2))), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false });
        Assert.Empty(result.Files); Assert.Equal(0, result.TotalBytes);
        Assert.Empty(Directory.GetFileSystemEntries(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }

    [Fact]
    public void ExcludedAliasActualSizeMustMatchEveryDeclaredFileRow() {
        using var fixture = new ArchiveFixtures();
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.MsiAliasFiles(false, new MsiFixtureFile("Template.adml"),
                new MsiFixtureFile("Template.adml", "hell", cabinet: 2, declaredSize: 5))), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.Contains("FileSize", error.Message);
        AssertNoOutput(fixture);
    }

    [Theory]
    [InlineData("Template.adml", "PolicyDefinitions/EN-US")]
    [InlineData("Témplate.adml", "PolicyDefinitions/en-US")]
    public void AliasCoalescingDoesNotPermitDirectoryCaseOrUnicodeNormalizationAmbiguity(string secondName, string secondParent) {
        using var fixture = new ArchiveFixtures();
        string firstName = secondName.StartsWith("Té", StringComparison.Ordinal) ? "Te\u0301mplate.adml" : "Template.adml";
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
            new MsiFixtureFile(firstName), new MsiFixtureFile(secondName, parent: secondParent))), fixture.Destination));
        AssertNoOutput(fixture);
    }

    [Fact]
    public void EveryAliasStillCountsTowardMemberExpansionAndCabinetLimits() {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.MsiAliasFiles(false,
            new MsiFixtureFile("Template.adml"), new MsiFixtureFile("Template.adml", cabinet: 2)));
        foreach (ArchiveExtractionOptions options in new[] {
            new ArchiveExtractionOptions { MaxEntryCount = 1 }, new ArchiveExtractionOptions { MaxExpandedBytes = 9 },
            new ArchiveExtractionOptions { MaxCabinetCount = 1 },
            new ArchiveExtractionOptions { MaxDecompressionWorkBytes = 32768 },
            new ArchiveExtractionOptions { IncludeFile = _ => false, MaxExpandedBytes = 9 }
        }) {
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination, options));
            AssertNoOutput(fixture);
        }
    }

    private static void AssertNoOutput(ArchiveFixtures fixture) {
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }
}
