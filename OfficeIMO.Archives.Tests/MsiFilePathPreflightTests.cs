namespace OfficeIMO.Archives.Tests;

public sealed class MsiFilePathPreflightTests {
#if !NETFRAMEWORK
    [Trait("Category", "Performance")]
#endif
    [Fact]
    public void ReusedOversizedFilenameIsRejectedWithoutAllocatingAnAssembledPathPerRow() {
        // Warm the public parser outside the measured operation. The package shares
        // one pool string across 2,000 rows; its input remains small while late path
        // validation would copy more than 32 MiB of rejected target strings.
        using (var warmup = new ArchiveFixtures())
            ArchiveExtractor.Extract(warmup.Package(ArchiveFixtures.Msi()), warmup.Destination);
        using var fixture = new ArchiveFixtures();
        string filename = new('x', 8192);
        var files = Enumerable.Range(0, 2000).Select(_ => new MsiFixtureFile(filename)).ToArray();
        string package = fixture.Package(ArchiveFixtures.MsiAliasFiles(false, files));
#if !NETFRAMEWORK
        long before = GC.GetAllocatedBytesForCurrentThread();
#endif
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(package, fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => false }));
#if !NETFRAMEWORK
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.True(allocated < 16 * 1024 * 1024,
            "Oversized shared filename must fail before repeated path copies; allocated " + allocated + " bytes.");
#endif
        Assert.False(Directory.Exists(fixture.Destination));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-archive-*"));
        Assert.Empty(Directory.GetDirectories(fixture.Root, ".officeimo-msi-alias-*"));
    }
}
