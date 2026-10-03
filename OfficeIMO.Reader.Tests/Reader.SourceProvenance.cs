using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderSourceProvenanceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Zip_root_describes_the_archive_and_preserves_member_provenance(bool empty) {
        byte[] bytes = BuildZip(empty);
        var reader = CreateZipReader();
        OfficeDocumentReadResult sync = reader.ReadDocument(bytes, "members.zip");
        OfficeDocumentReadResult asyncResult = await reader.ReadDocumentAsync(bytes, "members.zip");

        foreach (OfficeDocumentReadResult document in new[] { sync, asyncResult }) {
            Assert.Equal(ReaderInputKind.Zip, document.Kind);
            Assert.Equal("members.zip", document.Source.Path);
            Assert.Equal(bytes.Length, document.Source.LengthBytes);
            Assert.Equal(Sha256(bytes), document.Source.SourceHash);
            Assert.Contains("officeimo.reader.zip", document.CapabilitiesUsed);
            if (empty) continue;
            Assert.Equal(2, document.Chunks.Select(chunk => chunk.SourceId).Distinct().Count());
            Assert.All(document.Chunks, chunk => {
                Assert.NotEqual(document.Source.SourceId, chunk.SourceId);
                Assert.StartsWith("members.zip::", chunk.Location.Path);
                Assert.NotEqual(document.Source.SourceHash, chunk.SourceHash);
                Assert.True(chunk.SourceLengthBytes < document.Source.LengthBytes);
            });
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Async_file_reads_match_sync_identity_hashes_and_timestamps(bool empty) {
        string path = Path.Combine(Path.GetTempPath(), "reader-provenance-" + Guid.NewGuid().ToString("N") + ".txt");
        File.WriteAllText(path, empty ? string.Empty : "Alpha\nBeta");
        string relativePath = RelativePath(path);
        try {
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
            OfficeDocumentReadResult sync = reader.ReadDocument(relativePath);
            OfficeDocumentReadResult asyncResult = await reader.ReadDocumentAsync(relativePath);
            IReadOnlyList<ReaderChunk> asyncChunks = await reader.ReadAsync(relativePath);
            OfficeDocumentReadResult batch = Assert.Single(await reader.ReadDocumentsAsync(new[] { relativePath }));
            foreach (OfficeDocumentReadResult actual in new[] { asyncResult, batch }) {
                Assert.Equal(sync.Source.Path, actual.Source.Path);
                Assert.Equal(sync.Source.SourceId, actual.Source.SourceId);
                Assert.Equal(sync.Source.SourceHash, actual.Source.SourceHash);
                Assert.Equal(sync.Source.LastWriteUtc, actual.Source.LastWriteUtc);
                Assert.Equal(sync.Source.LengthBytes, actual.Source.LengthBytes);
                AssertChunkProvenance(sync.Chunks, actual.Chunks);
            }
            Assert.NotNull(sync.Source.SourceHash);
            Assert.NotNull(sync.Source.LastWriteUtc);
            AssertChunkProvenance(sync.Chunks, asyncChunks);
        } finally { File.Delete(path); }
    }

    [Fact]
    public async Task Async_archive_file_reads_preserve_member_identity() {
        string path = Path.Combine(Path.GetTempPath(), "reader-provenance-" + Guid.NewGuid().ToString("N") + ".zip");
        File.WriteAllBytes(path, BuildZip(false));
        string relativePath = RelativePath(path);
        try {
            var reader = CreateZipReader();
            OfficeDocumentReadResult sync = reader.ReadDocument(relativePath);
            OfficeDocumentReadResult asyncResult = await reader.ReadDocumentAsync(relativePath);
            Assert.Equal(sync.Source.SourceId, asyncResult.Source.SourceId);
            Assert.Equal(sync.Source.LastWriteUtc, asyncResult.Source.LastWriteUtc);
            AssertChunkProvenance(sync.Chunks, asyncResult.Chunks);
            AssertChunkProvenance(sync.Chunks, await reader.ReadAsync(relativePath));
        } finally { File.Delete(path); }
    }

    [Fact]
    public async Task Async_file_fallback_uses_the_registered_file_delegate() {
        string path = Path.Combine(Path.GetTempPath(), "reader-route-" + Guid.NewGuid().ToString("N") + ".rdr");
        File.WriteAllText(path, "input");
        try {
            var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
                Id = "file-specific", Kind = ReaderInputKind.Text, Extensions = new[] { ".rdr" },
                ReadPath = (name, _, _) => new[] { new ReaderChunk { Text = "file", Location = new ReaderLocation { Path = name } } },
                ReadStream = (_, name, _, _) => new[] { new ReaderChunk { Text = "stream", Location = new ReaderLocation { Path = name } } }
            }).Build();
            Assert.Equal("file", Assert.Single(await reader.ReadAsync(path)).Text);
            Assert.Equal("file", Assert.Single((await reader.ReadDocumentAsync(path)).Chunks).Text);
            Assert.Equal("stream", Assert.Single(reader.Read(Encoding.UTF8.GetBytes("input"), "input.rdr")).Text);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, false)]
    public async Task Processors_keep_member_provenance_when_local_ids_repeat(int copyMode, bool duplicateNames) {
        byte[] bytes = BuildZip(false, duplicateNames);
        OfficeDocumentReadResult original = CreateZipReader().ReadDocument(bytes, "members.zip");
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler()
            .AddProcessor(new DelegateOfficeDocumentProcessor("transform", (document, _) => {
                if (copyMode != 0) document.Chunks = document.Chunks.Select(chunk => new ReaderChunk {
                    Id = chunk.Id, Kind = chunk.Kind, Text = chunk.Text,
                    Location = copyMode == 1 ? chunk.Location : new ReaderLocation { Path = chunk.Location.Path },
                    SourceId = copyMode == 2 ? chunk.SourceId : null
                }).ToArray();
                foreach (ReaderChunk chunk in document.Chunks) chunk.Text += " processed";
                return document;
            })).Build();
        foreach (OfficeDocumentReadResult actual in new[] {
                     reader.ReadDocument(bytes, "members.zip"), await reader.ReadDocumentAsync(bytes, "members.zip") }) {
            Assert.Equal(original.Source.SourceHash, actual.Source.SourceHash);
            Assert.Equal(2, actual.Chunks.Select(chunk => chunk.SourceId).Distinct().Count());
            for (int index = 0; index < original.Chunks.Count; index++) {
                ReaderChunk before = original.Chunks[index];
                ReaderChunk after = actual.Chunks[index];
                Assert.Equal(before.SourceId, after.SourceId);
                Assert.Equal(before.SourceHash, after.SourceHash);
                Assert.Equal(before.SourceLengthBytes, after.SourceLengthBytes);
                Assert.Equal(before.SourceLastWriteUtc, after.SourceLastWriteUtc);
                Assert.NotEqual(before.ChunkHash, after.ChunkHash);
                Assert.EndsWith(" processed", after.Text);
            }
        }
    }

    [Fact]
    public void Folder_budget_charges_archive_bytes_instead_of_the_first_member() {
        string folder = Path.Combine(Path.GetTempPath(), "reader-budget-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        byte[] bytes = BuildZip(false);
        File.WriteAllBytes(Path.Combine(folder, "a.zip"), bytes);
        File.WriteAllBytes(Path.Combine(folder, "b.zip"), bytes);
        try {
            ReaderIngestResult result = CreateZipReader().ReadFolderDetailed(folder,
                new ReaderFolderOptions { Extensions = new[] { ".zip" }, MaxTotalBytes = bytes.Length });
            Assert.Equal(1, result.FilesParsed);
            Assert.Equal(1, result.FilesSkipped);
            Assert.Equal(bytes.Length, result.BytesRead);
            Assert.Equal(bytes.Length, result.Files[0].SourceLengthBytes);
        } finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public void Processor_root_does_not_inherit_a_handler_managed_child_hash() {
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "nested-managed", Kind = ReaderInputKind.Text, Extensions = new[] { ".rdr" },
            SourceHashBehavior = ReaderSourceHashBehavior.HandlerManaged,
            ReadStream = (_, name, _, _) => new[] {
                new ReaderChunk {
                    Id = "member", Text = "body", SourceId = "child-source", SourceHash = "child-hash",
                    SourceLengthBytes = 4, Location = new ReaderLocation { Path = name + "::child" }
                }
            }
        }).AddProcessor(new DelegateOfficeDocumentProcessor("pass", (document, _) => document)).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes("input"), "input.rdr",
            new ReaderOptions { ComputeHashes = false });
        Assert.Null(result.Source.SourceHash);
        Assert.Equal(5, result.Source.LengthBytes);
        Assert.Equal("child-source", Assert.Single(result.Chunks).SourceId);
        Assert.Equal("child-hash", result.Chunks[0].SourceHash);
        Assert.Equal(4, result.Chunks[0].SourceLengthBytes);
    }

    private static OfficeDocumentReader CreateZipReader() =>
        new OfficeDocumentReaderBuilder().AddPlainTextHandlers().AddZipHandler().Build();

    private static string RelativePath(string path) => Uri.UnescapeDataString(
        new Uri(Environment.CurrentDirectory + Path.DirectorySeparatorChar).MakeRelativeUri(new Uri(path)).ToString());

    private static byte[] BuildZip(bool empty, bool duplicateNames = false) {
        using var bytes = new MemoryStream();
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, true)) {
            if (!empty) foreach (string name in duplicateNames
                         ? new[] { "same.txt", "same.txt" } : new[] { "a/note.txt", "b/note.txt" }) {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open());
                writer.Write(name);
            }
        }
        return bytes.ToArray();
    }

    private static string Sha256(byte[] bytes) {
        using var hash = SHA256.Create();
        return string.Concat(hash.ComputeHash(bytes).Select(value => value.ToString("x2")));
    }

    private static void AssertChunkProvenance(IReadOnlyList<ReaderChunk> expected, IReadOnlyList<ReaderChunk> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++) {
            Assert.Equal(expected[index].SourceId, actual[index].SourceId);
            Assert.Equal(expected[index].SourceHash, actual[index].SourceHash);
            Assert.Equal(expected[index].SourceLastWriteUtc, actual[index].SourceLastWriteUtc);
            Assert.Equal(expected[index].SourceLengthBytes, actual[index].SourceLengthBytes);
            Assert.Equal(expected[index].ChunkHash, actual[index].ChunkHash);
        }
    }
}
