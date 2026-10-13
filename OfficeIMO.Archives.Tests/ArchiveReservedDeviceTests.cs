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
    [InlineData("CONIN$")]
    [InlineData("conout$")]
    [InlineData("CONIN$.txt")]
    [InlineData("conout$.tar.gz")]
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
    [InlineData("CONIN")]
    [InlineData("CONOUT")]
    [InlineData("CONIN$backup.txt")]
    [InlineData("xCONOUT$.txt")]
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

    [Theory]
    [InlineData("CONIN$", false, false)]
    [InlineData("CONOUT$", false, false)]
    [InlineData("conin$", false, true)]
    [InlineData("ConOut$", false, true)]
    [InlineData("CONIN$", true, false)]
    [InlineData("CONOUT$", true, false)]
    [InlineData("conin$", true, true)]
    [InlineData("ConOut$", true, true)]
    public void ConsoleDeviceNamesAreRejectedBeforeSelectionOrPublication(string name, bool msi, bool directory) {
        // Run the filtered case first: a regression fails before any device I/O is attempted.
        foreach (bool selected in new[] { false, true }) {
            using var fixture = new ArchiveFixtures();
            string authored = directory ? "templates/" + name + "/Template.admx" : "templates/" + name;
            byte[] package = msi
                ? ArchiveFixtures.MsiAliasFiles(false, directory
                    ? new MsiFixtureFile("Template.admx", parent: "templates/" + name)
                    : new MsiFixtureFile(name, parent: "templates"))
                : ArchiveFixtures.StoredCab(authored, new byte[] { 1 });
            int selectionCalls = 0;
            Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(package), fixture.Destination,
                new ArchiveExtractionOptions { IncludeFile = _ => { selectionCalls++; return selected; } }));
            Assert.Equal(0, selectionCalls);
            Assert.False(Directory.Exists(fixture.Destination));
            Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        }
    }

    private static ArchivePaths Paths() => new(new ArchiveLimits(new ArchiveExtractionOptions()));

}
