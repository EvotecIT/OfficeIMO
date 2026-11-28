using System.IO.Compression;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Archives;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveReaderFormatTests {
    [Fact]
    public void SevenZipMembersReachDocumentHandlers() {
        ReaderChunk[] chunks = Read(SevenZip("solid-lzma2.7z"), "notes.7z");
        Assert.Contains(chunks, chunk => chunk.Location.Path == "notes.7z::notes/café.txt" && chunk.Text.Contains("Harbour lights, chapter two."));
        Assert.Contains(chunks, chunk => chunk.Location.Path == "notes.7z::notes/second.txt" && chunk.Text.Contains("Harbour lights, chapter ten."));
        Assert.DoesNotContain(chunks, chunk => chunk.Kind == ReaderInputKind.Zip);
    }

    [Fact]
    public void CabinetMembersReachDocumentHandlers() {
        byte[] cabinet = ArchiveFixtures.StoredCab(("docs\\readme.txt", Encoding.UTF8.GetBytes("hello from a cabinet")));
        ReaderChunk chunk = Assert.Single(Read(cabinet, "policy.cab"));
        Assert.Equal("policy.cab::docs/readme.txt", chunk.Location.Path);
        Assert.Contains("hello from a cabinet", chunk.Text);
    }

    [Fact]
    public void CabinetFoldersAreDecodedOnceForAllTheirDocuments() {
        byte[] text = Encoding.UTF8.GetBytes(new string('a', 4000));
        byte[] cabinet = ArchiveFixtures.StoredCab(("one.txt", text), ("two.txt", text), ("three.txt", text));
        var budget = new ArchiveOperationBudget();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddArchiveHandler(new ArchiveTraversalOptions { OperationBudget = budget }).Build();
        ReaderChunk[] chunks = reader.Enumerate(new MemoryStream(cabinet), "three.cab", new ReaderOptions { ComputeHashes = false }).ToArray();
        Assert.Equal(3, chunks.Select(chunk => chunk.Location.Path).Distinct().Count());
        // One pass over the 12,000-byte folder, not one per member.
        Assert.Equal(3L * text.Length, budget.Usage.ExpandedBytes);
    }

    [Fact]
    public void CabinetMembersRespectTheTraversalExpansionLimits() {
        byte[] cabinet = ArchiveFixtures.StoredCab(("big.txt", Encoding.UTF8.GetBytes(new string('b', 2048))));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddArchiveHandler(new ArchiveTraversalOptions { MaxEntryUncompressedBytes = 1024 }).Build();
        Assert.Throws<ReaderResourceLimitException>(() => reader.Enumerate(new MemoryStream(cabinet), "big.cab").ToArray());

        reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddArchiveHandler(new ArchiveTraversalOptions { MaxTotalUncompressedBytes = 1024 }).Build();
        Assert.Throws<ReaderResourceLimitException>(() => reader.Enumerate(new MemoryStream(cabinet), "big.cab").ToArray());
    }

    [Fact]
    public void InstallerPackagesNeedTheOptIn() {
        byte[] msi = ArchiveFixtures.MsiFiles(false, false, 1, null, new MsiFixtureFile("readme.txt", "hello from an installer"));
        // Without the opt-in the handler does not claim .msi, so the package never reaches the archive reader.
        Assert.DoesNotContain(Read(msi, "setup.msi"), chunk => chunk.Kind == ReaderInputKind.Zip || (chunk.Location.Path ?? string.Empty).Contains("::"));

        ReaderChunk chunk = Assert.Single(Read(msi, "setup.msi", new ReaderArchiveOptions { ReadInstallerPackages = true }));
        Assert.StartsWith("setup.msi::", chunk.Location.Path);
        Assert.EndsWith("/readme.txt", chunk.Location.Path);
        Assert.Contains("hello from an installer", chunk.Text);
    }

    [Fact]
    public void NestedArchivesOfEveryKindShareOnePath() {
        byte[] outer = Zip(
            ("inner.7z", SevenZip("solid-lzma2.7z")),
            ("inner.cab", ArchiveFixtures.StoredCab(("readme.txt", Encoding.UTF8.GetBytes("cabinet inside zip")))),
            ("inner.tar.xz", TarWrapper("archive.tar.xz")),
            ("setup.msi", ArchiveFixtures.MsiFiles(false, false, 1, null, new MsiFixtureFile("readme.txt", "installer inside zip"))));

        ReaderChunk[] chunks = Read(outer, "outer.zip", new ReaderArchiveOptions { ReadInstallerPackages = true });

        Assert.Contains(chunks, chunk => chunk.Location.Path == "outer.zip::inner.7z::notes/café.txt");
        Assert.Contains(chunks, chunk => chunk.Location.Path == "outer.zip::inner.cab::readme.txt" && chunk.Text.Contains("cabinet inside zip"));
        Assert.Contains(chunks, chunk => chunk.Location.Path == "outer.zip::inner.tar.xz::docs/text.txt");
        Assert.Contains(chunks, chunk => (chunk.Location.Path ?? string.Empty).StartsWith("outer.zip::setup.msi::", StringComparison.Ordinal)
            && chunk.Text.Contains("installer inside zip"));

        // Without the opt-in, a nested installer is not an archive and not a document, so it is passed over.
        Assert.DoesNotContain(Read(outer, "outer.zip"), chunk => (chunk.Location.Path ?? string.Empty).Contains("setup.msi"));
    }

    [Fact]
    public void NestingDepthAndSizeLimitsApplyAcrossKinds() {
        byte[] innerZip = Zip(("deep.txt", Encoding.UTF8.GetBytes("three levels down")));
        byte[] cabinet = ArchiveFixtures.StoredCab(("inner.zip", innerZip));
        byte[] outer = Zip(("middle.cab", cabinet));

        Assert.Contains(Read(outer, "outer.zip", new ReaderArchiveOptions { MaxNestedDepth = 2 }),
            chunk => chunk.Location.Path == "outer.zip::middle.cab::inner.zip::deep.txt");

        ReaderChunk depth = Assert.Single(Read(outer, "outer.zip", new ReaderArchiveOptions { MaxNestedDepth = 1 }));
        Assert.Equal("outer.zip::middle.cab::inner.zip", depth.Location.Path);
        Assert.Contains("MaxNestedDepth (1)", depth.Text);

        ReaderChunk size = Assert.Single(Read(outer, "outer.zip", new ReaderArchiveOptions { MaxNestedArchiveBytes = 64 }));
        Assert.Equal("outer.zip::middle.cab", size.Location.Path);
        Assert.Contains("MaxNestedArchiveBytes", size.Text);

        ReaderChunk disabled = Assert.Single(Read(outer, "outer.zip", new ReaderArchiveOptions { ReadNestedArchives = false }));
        Assert.Contains("ReadNestedArchives is disabled", disabled.Text);
    }

    [Fact]
    public void DamagedNestedArchivesBecomeWarnings() {
        byte[] outer = Zip(("broken.7z", SevenZip("encrypted-member.7z")), ("ok.txt", Encoding.UTF8.GetBytes("still read")));
        ReaderChunk[] chunks = Read(outer, "outer.zip");
        Assert.Contains(chunks, chunk => chunk.Location.Path == "outer.zip::ok.txt" && chunk.Text.Contains("still read"));
        Assert.Contains(chunks, chunk => chunk.Location.Path == "outer.zip::broken.7z" && chunk.Text.Contains("Skipped nested archive"));
    }

    [Fact]
    public void RegisteredExtensionsFollowTheInstallerOption() {
        string[] plain = Extensions(new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddArchiveHandler().Build());
        Assert.Contains(".7z", plain);
        Assert.Contains(".cab", plain);
        Assert.DoesNotContain(".msi", plain);

        string[] withInstallers = Extensions(new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddArchiveHandler(readerArchiveOptions: new ReaderArchiveOptions { ReadInstallerPackages = true }).Build());
        Assert.Contains(".msi", withInstallers);
    }

    private static string[] Extensions(OfficeDocumentReader reader) =>
        reader.GetCapabilities().Single(capability => capability.Id == OfficeDocumentReaderBuilderArchiveExtensions.HandlerId).Extensions.ToArray();

    private static ReaderChunk[] Read(byte[] bytes, string name, ReaderArchiveOptions? options = null) {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddArchiveHandler(readerArchiveOptions: options).Build();
        return reader.Enumerate(new MemoryStream(bytes), name, new ReaderOptions { ComputeHashes = false }).ToArray();
    }

    private static byte[] Zip(params (string Name, byte[] Data)[] members) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var (name, data) in members) {
                using Stream entry = archive.CreateEntry(name).Open();
                entry.Write(data, 0, data.Length);
            }
        }
        return output.ToArray();
    }

    private static byte[] SevenZip(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentSevenZip", name));

    private static byte[] TarWrapper(string name) =>
        File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "IndependentTarWrappers", name));
}
