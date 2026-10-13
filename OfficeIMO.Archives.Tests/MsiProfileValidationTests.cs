using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class MsiProfileValidationTests {
    [Theory]
    [InlineData("OTHER", "SourceDir")]
    [InlineData("TARGETDIR", "Other")]
    public void DirectoryRootRequiresCanonicalInstallerProperties(string key, string definition) {
        using var fixture = new ArchiveFixtures();
        object?[][] rows = { new object?[] { key, null, definition } };
        byte[] package = ArchiveFixtures.Msi(directoryRows: rows, firstDirectory: key);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
            ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.Contains("TARGETDIR", error.Message);
        AssertNoOutput(fixture);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CanonicalRootAcceptsNullOrSelfParentWithoutAddingSourceDirToOutput(bool selfParent) {
        using var fixture = new ArchiveFixtures();
        object?[][] rows = { new object?[] { "TARGETDIR", selfParent ? "TARGETDIR" : null, "SourceDir" } };
        var result = ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Msi(directoryRows: rows,
            firstDirectory: "TARGETDIR")), fixture.Destination);
        Assert.Equal("Template.adml", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "Template.adml")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedSubstorageIsUnsupportedEvenWhenItContainsNoStreams(bool hasStream) {
        using var fixture = new ArchiveFixtures();
        Assert.True(OfficeCompoundFileReader.TryRead(ArchiveFixtures.Msi(), out OfficeCompoundFile? original, out string? error), error);
        Assert.NotNull(original);
        var entries = original.Entries.Concat(new[] { new OfficeCompoundFileEntry("ZCustomization", "ZCustomization", 1, 0) }).ToArray();
        var source = new OfficeCompoundFile(original.Streams, entries, original.RootEntry);
        var replacements = new Dictionary<string, byte[]>();
        if (hasStream) replacements.Add("ZCustomization/Metadata", new byte[] { 1, 2, 3 });
        byte[] package = OfficeCompoundFileWriter.Rewrite(source, replacements);
        NotSupportedException unsupported = Assert.Throws<NotSupportedException>(() =>
            ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination));
        Assert.Contains("substorage", unsupported.Message);
        AssertNoOutput(fixture);
    }

    private static void AssertNoOutput(ArchiveFixtures fixture) {
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }
}
