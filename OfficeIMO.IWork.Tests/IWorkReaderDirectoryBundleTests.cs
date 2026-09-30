using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkReaderDirectoryBundleTests {
    [Theory]
    [InlineData("nim-iwork/simple.pages", ".pages")]
    [InlineData("nim-iwork/simple.numbers", ".numbers")]
    [InlineData("nim-iwork/simple.key", ".key")]
    public async Task Bundle_is_one_document_with_snapshot_identity_for_sync_async_and_folder_reads(
        string fixture, string extension) {
        using var bundle = new ExtractedBundle(fixture, extension);
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        OfficeDocumentReadResult result = reader.ReadDocument(bundle.Path);
        OfficeDocumentReadResult asynchronous = await reader.ReadDocumentAsync(bundle.Path);
        IWorkSourceDocument source = IWorkSourceDocument.Open(bundle.Path);

        Assert.Equal(ReaderInputKind.IWork, result.Kind);
        Assert.Equal(source.ContainerLengthBytes, result.Source.LengthBytes);
        Assert.Equal(Directory.EnumerateFiles(bundle.Path, "*", SearchOption.AllDirectories)
            .Sum(path => new FileInfo(path).Length), result.Source.LengthBytes);
        Assert.Equal(IWorkSourceDocument.Open(Fixture(fixture)).ComputePackageContentHash(),
            result.Source.SourceHash);
        Assert.Equal(source.ComputePackageContentHash(), result.Source.SourceHash);
        Assert.Equal(result.Source.SourceHash, asynchronous.Source.SourceHash);
        Assert.Equal(result.Chunks.Select(chunk => chunk.Text), asynchronous.Chunks.Select(chunk => chunk.Text));
        Assert.All(result.Chunks, chunk => {
            Assert.Equal(bundle.Path, chunk.Location.Path);
            Assert.Equal(result.Source.SourceHash, chunk.SourceHash);
            Assert.Equal(result.Source.LengthBytes, chunk.SourceLengthBytes);
        });
        Assert.Equal(result.Chunks.Count, (await reader.ReadAsync(bundle.Path)).Count);
        Assert.True(Assert.Single(reader.GetCapabilities()).SupportsDirectoryBundle);
        Assert.Equal(ReaderInputKind.IWork, reader.Detect(bundle.Path).Kind);
        Assert.False(reader.Detect(bundle.Path).ContentInspected);
        string trailing = bundle.Path + System.IO.Path.DirectorySeparatorChar;
        Assert.Equal(result.Source.SourceHash, reader.ReadDocument(trailing).Source.SourceHash);
        Assert.Equal(result.Source.SourceHash, (await reader.ReadDocumentAsync(trailing)).Source.SourceHash);
        Assert.Equal(ReaderInputKind.IWork, reader.Detect(trailing).Kind);
        Assert.Equal(new[] { bundle.Path }, reader.EnumerateDocumentPaths(new[] { trailing }).ToArray());
        Assert.Equal(1, reader.ReadFolderDetailed(trailing).FilesParsed);

        ReaderPathDocumentResult detailed = reader.ReadPathDocumentsDetailed(bundle.Path);
        Assert.Equal(1, detailed.FilesParsed);
        Assert.Equal(result.Source.LengthBytes, detailed.BytesRead);
        ReaderIngestResult folder = reader.ReadFolderDetailed(bundle.Root);
        Assert.Equal(1, folder.FilesParsed);
        Assert.Equal(result.Source.SourceHash, Assert.Single(folder.Files).SourceHash);
        Assert.Equal(result.Source.LengthBytes, folder.BytesRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Folder_budget_preserves_the_selected_handlers_default_input_limit(bool preferContent) {
        using var bundle = new ExtractedBundle("nim-iwork/simple.pages", ".pages");
        string path = System.IO.Path.Combine(bundle.Root, preferContent ? "sample.txt" : "sample.bounded");
        File.WriteAllText(path, preferContent ? "# Detected Markdown\n\nBody" : "12345678901");
        int calls = 0;
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "bounded-test", Kind = ReaderInputKind.Text, Extensions = new[] { ".bounded", ".txt" },
            DefaultMaxInputBytes = preferContent ? 100 : 10,
            ReadPath = (_, _, _) => { calls++; return Array.Empty<ReaderChunk>(); }
        }).AddHandler(new ReaderHandlerRegistration {
            Id = "markdown-test", Kind = ReaderInputKind.Markdown, Extensions = new[] { ".md" },
            DefaultMaxInputBytes = 10,
            ReadPath = (_, _, _) => { calls++; return Array.Empty<ReaderChunk>(); }
        }).Build();
        var options = new ReaderOptions { DetectionMode = preferContent ? ReaderDetectionMode.PreferContent : ReaderDetectionMode.ContentWhenUnknown };
        Assert.Throws<IOException>(() => reader.ReadDocument(path, options));
        ReaderIngestResult folder = reader.ReadFolderDetailed(bundle.Root,
            new ReaderFolderOptions { MaxTotalBytes = 100 }, options);
        Assert.Equal(0, folder.FilesParsed);
        Assert.Equal(1, folder.FilesSkipped);
        Assert.Equal(0, calls);
    }

    [Fact]
    public void Folder_discovery_does_not_descend_into_a_registered_bundle() {
        using var bundle = new ExtractedBundle("nim-iwork/simple.pages", ".pages");
        File.Copy(Fixture("nim-iwork/simple.pages"), System.IO.Path.Combine(bundle.Path, "inside.pages"));
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        Assert.Equal(new[] { bundle.Path }, reader.EnumerateDocumentPaths(new[] { bundle.Root }).ToArray());
        Assert.Equal(new[] { bundle.Path }, reader.EnumerateDocumentPaths(new[] { bundle.Path }).ToArray());
        Assert.Equal(1, reader.ReadFolderDetailed(bundle.Root).FilesParsed);
        Assert.Equal(1, reader.ReadFolderDetailed(bundle.Path).FilesParsed);
    }

    [Fact]
    public void Bundle_input_and_folder_byte_limits_are_enforced_by_the_package_owner() {
        using var bundle = new ExtractedBundle("nim-iwork/simple.pages", ".pages");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        long length = IWorkSourceDocument.Open(bundle.Path).ContainerLengthBytes;
        Assert.Throws<InvalidDataException>(() => reader.ReadDocument(bundle.Path,
            new ReaderOptions { MaxInputBytes = length - 1 }));
        ReaderIngestResult skipped = reader.ReadFolderDetailed(bundle.Root,
            new ReaderFolderOptions { MaxTotalBytes = length - 1 });
        Assert.Equal(0, skipped.FilesParsed);
        Assert.Equal(1, skipped.FilesSkipped);
        Assert.Empty(reader.EnumerateDocumentPaths(new[] { bundle.Root },
            new ReaderFolderOptions { MaxTotalBytes = length - 1 }));
        Assert.Empty(reader.EnumerateDocumentPaths(new[] { bundle.Path },
            new ReaderFolderOptions { MaxTotalBytes = length - 1 }));
        Assert.Equal(new[] { bundle.Path }, reader.EnumerateDocumentPaths(new[] { bundle.Root },
            new ReaderFolderOptions { MaxTotalBytes = length }).ToArray());
        Assert.Equal(length, reader.ReadDocument(bundle.Path,
            new ReaderOptions { MaxInputBytes = length }).Source.LengthBytes);
    }

    [Fact]
    public void Snapshot_hash_changes_with_entry_name_or_bytes_and_can_be_disabled() {
        using var bundle = new ExtractedBundle("nim-iwork/simple.pages", ".pages");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        string first = reader.ReadDocument(bundle.Path).Source.SourceHash!;
        string resource = System.IO.Path.Combine(bundle.Path, "resource.bin");
        File.WriteAllBytes(resource, new byte[] { 1 });
        string added = reader.ReadDocument(bundle.Path).Source.SourceHash!;
        File.WriteAllBytes(resource, new byte[] { 2 });
        string changed = reader.ReadDocument(bundle.Path).Source.SourceHash!;
        File.Move(resource, System.IO.Path.Combine(bundle.Path, "renamed.bin"));
        string renamed = reader.ReadDocument(bundle.Path).Source.SourceHash!;
        Assert.NotEqual(first, added);
        Assert.NotEqual(added, changed);
        Assert.NotEqual(changed, renamed);
        Assert.Null(reader.ReadDocument(bundle.Path, new ReaderOptions { ComputeHashes = false }).Source.SourceHash);
    }

    [Fact]
    public async Task Bundle_cancellation_and_invalid_package_do_not_fall_back_to_folder_ingestion() {
        using var bundle = new ExtractedBundle("nim-iwork/simple.pages", ".pages");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => reader.ReadDocument(bundle.Path,
            cancellationToken: cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadDocumentAsync(bundle.Path,
            cancellationToken: cancellation.Token));
        string invalid = System.IO.Path.Combine(bundle.Root, "empty.pages");
        Directory.CreateDirectory(invalid);
        Assert.Throws<InvalidDataException>(() => reader.ReadDocument(invalid));
        OfficeDocumentReader unregistered = new OfficeDocumentReaderBuilder().Build();
        Assert.Throws<IOException>(() => unregistered.ReadDocument(bundle.Path));
    }

    private static string Fixture(string name) => System.IO.Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", name);

    private sealed class ExtractedBundle : IDisposable {
        internal ExtractedBundle(string fixture, string extension) {
            Root = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-reader-bundle-" + Guid.NewGuid().ToString("N"));
            Path = System.IO.Path.Combine(Root, "document" + extension);
            Directory.CreateDirectory(Path);
            ZipFile.ExtractToDirectory(Fixture(fixture), Path);
        }
        internal string Root { get; }
        internal string Path { get; }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
