#if NET8_0_OR_GREATER
namespace OfficeIMO.Archives.Tests;

/// <summary>Protects extraction beneath normal Unix aliases and the controlled staging link boundary.</summary>
public sealed class ArchiveSymlinkTests {
    [UnixSymlinkFact]
    public void CallerSymlinkedAncestorCanHostPackageAndDestination() {
        using var fixture = new ArchiveFixtures();
        string physical = Path.Combine(fixture.Root, "physical");
        Directory.CreateDirectory(physical);
        string alias = Path.Combine(fixture.Root, "alias");
        Directory.CreateSymbolicLink(alias, physical);
        File.WriteAllBytes(Path.Combine(physical, "package.cab"), ArchiveFixtures.StoredCab("nested/file.txt", Encoding.UTF8.GetBytes("hello")));

        var result = ArchiveExtractor.Extract(Path.Combine(alias, "package.cab"), Path.Combine(alias, "output"));

        Assert.Equal("nested/file.txt", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(physical, "output", "nested", "file.txt")));
        Assert.Empty(Directory.GetDirectories(physical, ".officeimo-archive-*"));
    }

    [UnixSymlinkFact]
    public void ReadOnlyPackageSymlinkCanBeOpened() {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab("file.txt", Encoding.UTF8.GetBytes("hello")));
        string alias = Path.Combine(fixture.Root, "linked-package.cab");
        File.CreateSymbolicLink(alias, package);

        var result = ArchiveExtractor.Extract(alias, fixture.Destination);

        Assert.Equal("file.txt", Assert.Single(result.Files).RelativePath);
        Assert.Equal("hello", File.ReadAllText(Path.Combine(fixture.Destination, "file.txt")));
    }

    [UnixSymlinkFact]
    public void DestinationLinkIsRejectedWithoutChangingItsTarget() {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab("file.txt", new byte[] { 1 }));
        string outside = Path.Combine(fixture.Root, "outside");
        Directory.CreateDirectory(outside);
        Directory.CreateSymbolicLink(fixture.Destination, outside);

        Assert.Throws<IOException>(() => ArchiveExtractor.Extract(package, fixture.Destination));

        Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
        Assert.NotNull(new DirectoryInfo(fixture.Destination).LinkTarget);
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    [UnixSymlinkTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void LinksIntroducedInsideStagingCannotCreateDirectoriesOutsideIt(bool replaceStagingRoot) {
        using var fixture = new ArchiveFixtures();
        string package = fixture.Package(ArchiveFixtures.StoredCab("nested/new/file.txt", new byte[] { 1 }));
        string outside = Path.Combine(fixture.Root, "outside");
        Directory.CreateDirectory(outside);
        var options = new ArchiveExtractionOptions {
            IncludeFile = _ => {
                string staging = Assert.Single(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
                if (replaceStagingRoot) {
                    Directory.Delete(staging);
                    Directory.CreateSymbolicLink(staging, outside);
                } else {
                    Directory.CreateSymbolicLink(Path.Combine(staging, "nested"), outside);
                }
                return true;
            }
        };

        Assert.Throws<IOException>(() => ArchiveExtractor.Extract(package, fixture.Destination, options));

        Assert.Empty(Directory.EnumerateFileSystemEntries(outside));
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    private sealed class UnixSymlinkFactAttribute : FactAttribute {
        public UnixSymlinkFactAttribute() {
            if (OperatingSystem.IsWindows()) Skip = "Exercises native Unix symbolic links without requiring Windows link privileges.";
        }
    }

    private sealed class UnixSymlinkTheoryAttribute : TheoryAttribute {
        public UnixSymlinkTheoryAttribute() {
            if (OperatingSystem.IsWindows()) Skip = "Exercises native Unix symbolic links without requiring Windows link privileges.";
        }
    }
}
#endif
