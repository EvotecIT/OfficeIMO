namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveExtractorTests {
    [Fact]
    public void CabSelectionPreservesAuthoredCaseAndTreatsNestedArchivesAsFiles() {
        using var fixture = new ArchiveFixtures();
        byte[] inner = ArchiveFixtures.StoredCab("ignored.txt", Encoding.UTF8.GetBytes("inner"));
        string package = fixture.Package(ArchiveFixtures.StoredCab(("Windows/ADMX/en-US/Example.adml", Encoding.UTF8.GetBytes("hello")), ("nested.cab", inner)));
        ArchiveExtractionResult result = ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = path => path.EndsWith(".adml", StringComparison.Ordinal) });
        Assert.Equal(ArchiveFormat.Cab, result.Format);
        Assert.Equal("Windows/ADMX/en-US/Example.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal(5, result.TotalBytes);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "Windows", "ADMX", "en-US", "Example.adml")));
    }
    [Theory]
    [InlineData("../escape.txt")]
    [InlineData("C:/escape.txt")]
    [InlineData("file:stream")]
    public void UnsafeCabNamesFailBeforePublishingEvenWhenFilteredOut(string name) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab(name, new byte[] { 1 }));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }
    [Theory]
    [InlineData("file.txt", "FILE.txt")]
    [InlineData("folder", "folder/file.txt")]
    [InlineData("Folder/a.txt", "folder/b.txt")]
    public void PortableCollisionsRejectTheWholeResult(string first, string second) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab((first, new byte[] { 1 }), (second, new byte[] { 2 })));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }
    [Fact]
    public void LimitsAndCancellationKeepExistingDestinationContents() {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab("file.txt", Encoding.UTF8.GetBytes("hello")));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination, new ArchiveExtractionOptions { MaxExpandedBytes = 4 }));
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => ArchiveExtractor.Extract(package, fixture.Destination, cancellationToken: canceled.Token));
        Directory.CreateDirectory(fixture.Destination); File.WriteAllText(Path.Combine(fixture.Destination, "keep.txt"), "keep");
        Assert.Throws<IOException>(() => ArchiveExtractor.Extract(package, fixture.Destination));
        Assert.Equal("keep", File.ReadAllText(Path.Combine(fixture.Destination, "keep.txt")));
    }
    [Fact]
    public void StoredCabExtractionUsesTheSamePathAndPublicationContract() {
        using var fixture = new ArchiveFixtures();
        ArchiveExtractionResult result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.StoredCab("locale/Example.adml", Encoding.UTF8.GetBytes("hello"))), fixture.Destination);
        Assert.Equal(ArchiveFormat.Cab, result.Format);
        Assert.Equal("locale/Example.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "locale", "Example.adml")));
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void MsiUsesAuthoredPathsAndDeclaredSchemasWithoutAnInstaller(bool wideMedia, bool longStrings) {
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(wideMedia, longStrings)), fixture.Destination);
        Assert.Equal(ArchiveFormat.Msi, result.Format);
        Assert.Equal("PolicyDefinitions/en-US/Template.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "en-US", "Template.adml")));
    }
    [Fact]
    public void MissingOrExternalMsiCabinetAndDirectoryCyclesCannotPublishPartialResults() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(externalCab: true)), fixture.Destination));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(missingCabinet: true)), fixture.Destination));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(directoryCycle: true)), fixture.Destination));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(wrongSize: true)), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void UnreferencedDirectoryChainsCannotBypassStructuralDepthThroughCachedParents(bool childFirst, bool dots) {
        using var fixture = new ArchiveFixtures();
        var directories = new List<object?[]> { new object?[] { "TARGETDIR", null, "SourceDir" } };
        for (int index = 1; index <= 20; index++) directories.Add(new object?[] { "D" + index, index == 1 ? "TARGETDIR" : "D" + (index - 1), dots ? "." : "folder" });
        if (childFirst) directories.Reverse();
        string package = fixture.Package(ArchiveFixtures.Msi(directoryRows: directories.ToArray(), firstDirectory: "TARGETDIR"));
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { MaxPathDepth = 8 }));
        Assert.False(Directory.Exists(fixture.Destination));
    }
    [Fact]
    public void UnreferencedDirectoryPrefixesObservePathBoundsBeforeCaching() {
        using var fixture = new ArchiveFixtures();
        object?[][] directories = { new object?[] { "TARGETDIR", null, "SourceDir" },
            new object?[] { "PARENT", "TARGETDIR", "12345678901234567890" },
            new object?[] { "CHILD", "PARENT", "12345678901234567890" } };
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(directoryRows: directories, firstDirectory: "TARGETDIR")), fixture.Destination,
            new ArchiveExtractionOptions { MaxPathLength = 32 }));
        Assert.False(Directory.Exists(fixture.Destination));
    }
    [Fact]
    public void DistinctCaseSensitiveDirectoryIdentifiersCanMapToDistinctOutputFolders() {
        using var fixture = new ArchiveFixtures();
        object?[][] directories = { new object?[] { "TARGETDIR", null, "SourceDir" },
            new object?[] { "Data", "TARGETDIR", "Upper" }, new object?[] { "data", "TARGETDIR", "Lower" } };
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(directoryRows: directories,
            firstDirectory: "Data", secondDirectory: "data", twoCabinets: true)), fixture.Destination);
        Assert.Equal(new[] { "Lower/Second.adml", "Upper/Template.adml" }, result.Files.Select(file => file.RelativePath));
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "Upper", "Template.adml")));
        Assert.Equal("world", File.ReadAllText(Path.Combine(fixture.Destination, "Lower", "Second.adml")));
    }
    [Fact]
    public void DirectoryForeignKeysRequireTheAuthoredCase() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(firstDirectory: "locale")), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }
    [Fact]
    public void MultipleCabinetsUseDiskOrderAndRejectBackwardSequenceRanges() {
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)), fixture.Destination);
        Assert.Equal(2, result.Files.Count);
        Assert.Equal("world", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "en-US", "Second.adml")));
        string invalidDestination = Path.Combine(fixture.Root, "invalid");
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(twoCabinets: true, backwardMedia: true)), invalidDestination));
        Assert.False(Directory.Exists(invalidDestination));
    }
    [Fact]
    public void FileIdentifiersRemainCaseInsensitiveWhenMatchingCabinetMembers() {
        using var fixture = new ArchiveFixtures();
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(lowercaseCabinetKey: true)), fixture.Destination);
        Assert.Equal("PolicyDefinitions/en-US/Template.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "PolicyDefinitions", "en-US", "Template.adml")));
    }
    [Fact]
    public void FilteredMsiCabinetsStillShareTheAggregateDecodedByteBudget() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.Msi(twoCabinets: true)), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false, MaxExpandedBytes = 9 }));
        Assert.False(Directory.Exists(fixture.Destination));
    }
}
