namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveReservedDeviceTests {
    [Theory]
    [InlineData("COM\u00b9.txt")]
    [InlineData("com\u00b2.tar.gz")]
    [InlineData("CoM\u00b3")]
    [InlineData("LPT\u00b9")]
    [InlineData("lpt\u00b2.log")]
    [InlineData("LpT\u00b3.any")]
    [InlineData("COM9.txt")]
    [InlineData("lpt1.tar.gz")]
    public void WindowsDeviceNamesAreRejectedAsFilesAndDirectorySegments(string name) {
        Assert.Throws<InvalidDataException>(() => Paths().Add(name));
        Assert.Throws<InvalidDataException>(() => Paths().Add("root/" + name + "/template.admx"));
    }

    [Theory]
    [InlineData("COM0")]
    [InlineData("LPT10.log")]
    [InlineData("COM\u2074.txt")]
    [InlineData("xCOM\u00b9.txt")]
    [InlineData("COM\u00b9data.txt")]
    [InlineData("LPT\u00b2backup")]
    [InlineData("BOX\u00b9.txt")]
    [InlineData("R\u00e9sum\u00e9.adml")]
    public void OrdinaryNearNamesAndUnicodePreserveTheirAuthoredSpelling(string name) {
        string path = "templates/" + name;
        Assert.Equal(path, Paths().Add(path));
    }

    [Theory]
    [InlineData("templates/COM\u00b9.txt")]
    [InlineData("LPT\u00b3/Template.admx")]
    public void Utf8CabinetDevicePathsAreRejectedBeforeSelectionOrPublication(string name) {
        using var fixture = new ArchiveFixtures();
        int selectionCalls = 0;
        // The predicate excludes every file, so the RED run never attempts device I/O.
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.Utf8Cabinet(name)), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return false; } }));
        Assert.Equal(0, selectionCalls);
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
    }

    private static ArchivePaths Paths() => new(new ArchiveLimits(new ArchiveExtractionOptions()));

}
