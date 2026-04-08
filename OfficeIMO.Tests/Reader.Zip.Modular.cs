using OfficeIMO.Reader;
using OfficeIMO.Reader.Zip;
using OfficeIMO.Zip;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderZipModularTests {
    [Fact]
    public void ZipTraversal_EmitsWarningsAndRespectsLimits() {
        var zipPath = Path.Combine(Path.GetTempPath(), "officeimo-zip-" + Guid.NewGuid().ToString("N") + ".zip");
        try {
            using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false)) {
                WriteTextEntry(archive, "safe/one.txt", "one");
                WriteTextEntry(archive, "safe/two.txt", "two");
                WriteTextEntry(archive, "../evil.txt", "bad");
                WriteTextEntry(archive, "/absolute.txt", "bad");
                WriteTextEntry(archive, "deep/a/b/c/d.txt", "deep");
            }

            var result = ZipTraversal.Traverse(zipPath, new ZipTraversalOptions {
                MaxEntries = 2,
                MaxDepth = 3,
                MaxTotalUncompressedBytes = 128,
                MaxEntryUncompressedBytes = 32,
                DeterministicOrder = true
            });

            Assert.NotNull(result);
            Assert.True(result.Entries.Count <= 2);
            Assert.All(result.Entries, e => Assert.True(e.Depth <= 3));
            Assert.All(result.Entries, e => Assert.DoesNotContain("..", e.FullName, StringComparison.Ordinal));
            Assert.Contains(result.Warnings, w => w.Warning.Contains("path traversal", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(result.Warnings, w => w.Warning.Contains("MaxDepth", StringComparison.OrdinalIgnoreCase) || w.Warning.Contains("MaxEntries", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (File.Exists(zipPath)) File.Delete(zipPath);
        }
    }

    [Fact]
    public void ZipTraversal_RespectsCompressionRatioLimit() {
        var zipPath = Path.Combine(Path.GetTempPath(), "officeimo-zip-" + Guid.NewGuid().ToString("N") + ".zip");
        try {
            using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false)) {
                WriteTextEntry(archive, "bomb.txt", new string('A', 25_000));
            }

            var result = ZipTraversal.Traverse(zipPath, new ZipTraversalOptions {
                MaxCompressionRatio = 2
            });

            Assert.Empty(result.Entries);
            Assert.Contains(result.Warnings, w => w.Warning.Contains("MaxCompressionRatio", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (File.Exists(zipPath)) File.Delete(zipPath);
        }
    }

    [Fact]
    public void DocumentReaderZip_ReadsNestedZipAndEmitsWarnings() {
        var zipPath = Path.Combine(Path.GetTempPath(), "officeimo-reader-zip-" + Guid.NewGuid().ToString("N") + ".zip");
        try {
            var nestedZipBytes = BuildNestedZipBytes();

            using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false)) {
                WriteTextEntry(archive, "docs/readme.md", "# Top\n\nHello from top");
                WriteBytesEntry(archive, "nested/archive.zip", nestedZipBytes);
                WriteTextEntry(archive, "big.txt", new string('x', 2048));
            }

            var chunks = DocumentReaderZipExtensions.ReadZip(
                zipPath,
                readerOptions: new ReaderOptions { MaxInputBytes = 512, MaxChars = 8_000 },
                zipOptions: new ZipTraversalOptions { DeterministicOrder = true },
                readerZipOptions: new ReaderZipOptions { ReadNestedZipEntries = true, MaxNestedDepth = 2 }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("docs/readme.md", StringComparison.OrdinalIgnoreCase) ?? false));

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("nested/archive.zip::nested.md", StringComparison.OrdinalIgnoreCase) ?? false));

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Zip &&
                (c.Warnings?.Any(w => w.Contains("MaxInputBytes", StringComparison.OrdinalIgnoreCase)) ?? false));
        } finally {
            if (File.Exists(zipPath)) File.Delete(zipPath);
        }
    }

    [Fact]
    public void DocumentReaderZip_ReadsFromNonSeekableStream() {
        var zipBytes = BuildSimpleZipBytes();
        using var stream = new NonSeekableReadStream(zipBytes);

        var chunks = DocumentReaderZipExtensions.ReadZip(
            stream,
            sourceName: "nonseekable.zip",
            readerOptions: new ReaderOptions { MaxChars = 8_000 },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true }).ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("nonseekable.zip::docs/readme.md", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("From non-seekable stream.", StringComparison.Ordinal) ?? false));
    }

    [Fact]
    public void DocumentReaderZip_ReadsFromNonSeekableStream_EnforcesMaxInputBytes() {
        var zipBytes = BuildSimpleZipBytes();
        using var stream = new NonSeekableReadStream(zipBytes);

        var ex = Assert.Throws<IOException>(() => DocumentReaderZipExtensions.ReadZip(
            stream,
            sourceName: "nonseekable.zip",
            readerOptions: new ReaderOptions { MaxInputBytes = 16, MaxChars = 8_000 },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderZip_UsesVirtualPathForSourceIdentity() {
        var zipBytes = BuildSimpleZipBytes();

        using var first = new MemoryStream(zipBytes, writable: false);
        using var second = new MemoryStream(zipBytes, writable: false);

        var firstChunk = DocumentReaderZipExtensions.ReadZip(
            first,
            sourceName: "first.zip",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true })
            .Single(c => c.Kind == ReaderInputKind.Markdown);

        var secondChunk = DocumentReaderZipExtensions.ReadZip(
            second,
            sourceName: "second.zip",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true })
            .Single(c => c.Kind == ReaderInputKind.Markdown);

        Assert.Contains("first.zip::docs/readme.md", firstChunk.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("second.zip::docs/readme.md", secondChunk.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.False(string.IsNullOrWhiteSpace(firstChunk.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(secondChunk.SourceHash));
        Assert.True(firstChunk.TokenEstimate.HasValue && firstChunk.TokenEstimate.Value >= 1);
        Assert.True(secondChunk.TokenEstimate.HasValue && secondChunk.TokenEstimate.Value >= 1);
        Assert.NotEqual(firstChunk.SourceId, secondChunk.SourceId);
        Assert.NotEqual(firstChunk.ChunkHash, secondChunk.ChunkHash);
        Assert.Equal(firstChunk.SourceHash, secondChunk.SourceHash);
    }

    [Fact]
    public void DocumentReaderZip_WarningChunks_UseVirtualPathForSourceIdentity() {
        var zipBytes = BuildZipWithLargeEntryBytes();

        using var first = new MemoryStream(zipBytes, writable: false);
        using var second = new MemoryStream(zipBytes, writable: false);

        var firstWarning = DocumentReaderZipExtensions.ReadZip(
            first,
            sourceName: "first.zip",
            readerOptions: new ReaderOptions { MaxInputBytes = 512, MaxChars = 8_000, ComputeHashes = true },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true })
            .Single(c => c.Kind == ReaderInputKind.Zip && (c.Warnings?.Count ?? 0) > 0);

        var secondWarning = DocumentReaderZipExtensions.ReadZip(
            second,
            sourceName: "second.zip",
            readerOptions: new ReaderOptions { MaxInputBytes = 512, MaxChars = 8_000, ComputeHashes = true },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true })
            .Single(c => c.Kind == ReaderInputKind.Zip && (c.Warnings?.Count ?? 0) > 0);

        Assert.Contains("first.zip::big.txt", firstWarning.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("second.zip::big.txt", secondWarning.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(firstWarning.SourceId);
        Assert.NotNull(secondWarning.SourceId);
        Assert.False(string.IsNullOrWhiteSpace(firstWarning.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(secondWarning.SourceHash));
        Assert.NotEqual(firstWarning.SourceId, secondWarning.SourceId);
        Assert.NotNull(firstWarning.ChunkHash);
        Assert.NotNull(secondWarning.ChunkHash);
        Assert.True(firstWarning.TokenEstimate.HasValue && firstWarning.TokenEstimate.Value >= 1);
        Assert.True(secondWarning.TokenEstimate.HasValue && secondWarning.TokenEstimate.Value >= 1);
        Assert.NotEqual(firstWarning.ChunkHash, secondWarning.ChunkHash);
        Assert.Equal(2048, firstWarning.SourceLengthBytes);
        Assert.Equal(2048, secondWarning.SourceLengthBytes);
    }

    [Fact]
    public void DocumentReaderZip_PreservesCaseSensitiveVirtualSourceIdentity() {
        var zipBytes = BuildZipWithCaseDistinctEntries();

        using var stream = new MemoryStream(zipBytes, writable: false);

        var chunks = DocumentReaderZipExtensions.ReadZip(
            stream,
            sourceName: "case-sensitive.zip",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true },
            zipOptions: new ZipTraversalOptions { DeterministicOrder = true })
            .Where(c => c.Kind == ReaderInputKind.Markdown)
            .ToList();

        var upperChunk = Assert.Single(chunks, c => string.Equals(c.Location.Path, "case-sensitive.zip::Docs/Readme.md", StringComparison.Ordinal));
        var lowerChunk = Assert.Single(chunks, c => string.Equals(c.Location.Path, "case-sensitive.zip::docs/readme.md", StringComparison.Ordinal));

        Assert.NotEqual(upperChunk.SourceId, lowerChunk.SourceId);
        Assert.NotEqual(upperChunk.ChunkHash, lowerChunk.ChunkHash);
        Assert.Equal(upperChunk.SourceHash, lowerChunk.SourceHash);
    }

    [Fact]
    public void DocumentReaderZip_DisabledNestedZipWarnings_IncludeEntryMetadata() {
        var zipPath = Path.Combine(Path.GetTempPath(), "officeimo-reader-zip-disabled-nested-" + Guid.NewGuid().ToString("N") + ".zip");
        try {
            var nestedZipBytes = BuildNestedZipBytes();

            using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false)) {
                WriteBytesEntry(archive, "nested/archive.zip", nestedZipBytes);
            }

            var warningChunk = DocumentReaderZipExtensions.ReadZip(
                zipPath,
                readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true },
                zipOptions: new ZipTraversalOptions { DeterministicOrder = true },
                readerZipOptions: new ReaderZipOptions { ReadNestedZipEntries = false })
                .Single(c => c.Kind == ReaderInputKind.Zip && (c.Warnings?.Any(w => w.Contains("ReadNestedZipEntries is disabled", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.Contains("nested/archive.zip", warningChunk.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            Assert.NotNull(warningChunk.SourceId);
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceHash));
            Assert.NotNull(warningChunk.ChunkHash);
            Assert.True(warningChunk.TokenEstimate.HasValue && warningChunk.TokenEstimate.Value >= 1);
            Assert.True(warningChunk.SourceLengthBytes > 0);
            Assert.NotNull(warningChunk.SourceLastWriteUtc);
        } finally {
            if (File.Exists(zipPath)) File.Delete(zipPath);
        }
    }

    [Fact]
    public void DocumentReaderZip_MaxNestedDepthWarnings_IncludeEntryMetadata() {
        var zipPath = Path.Combine(Path.GetTempPath(), "officeimo-reader-zip-max-depth-" + Guid.NewGuid().ToString("N") + ".zip");
        try {
            var nestedZipBytes = BuildNestedZipBytes();

            using (var fs = new FileStream(zipPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fs, ZipArchiveMode.Create, leaveOpen: false)) {
                WriteBytesEntry(archive, "nested/archive.zip", nestedZipBytes);
            }

            var warningChunk = DocumentReaderZipExtensions.ReadZip(
                zipPath,
                readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true },
                zipOptions: new ZipTraversalOptions { DeterministicOrder = true },
                readerZipOptions: new ReaderZipOptions { ReadNestedZipEntries = true, MaxNestedDepth = 0 })
                .Single(c => c.Kind == ReaderInputKind.Zip && (c.Warnings?.Any(w => w.Contains("MaxNestedDepth", StringComparison.OrdinalIgnoreCase)) ?? false));

            Assert.Contains("nested/archive.zip", warningChunk.Location.Path ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            Assert.NotNull(warningChunk.SourceId);
            Assert.False(string.IsNullOrWhiteSpace(warningChunk.SourceHash));
            Assert.NotNull(warningChunk.ChunkHash);
            Assert.True(warningChunk.TokenEstimate.HasValue && warningChunk.TokenEstimate.Value >= 1);
            Assert.True(warningChunk.SourceLengthBytes > 0);
            Assert.NotNull(warningChunk.SourceLastWriteUtc);
        } finally {
            if (File.Exists(zipPath)) File.Delete(zipPath);
        }
    }

    private static byte[] BuildNestedZipBytes() {
        using var ms = new MemoryStream();
        using (var nestedArchive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(nestedArchive, "nested.md", "# Nested\n\nBody");
        }

        return ms.ToArray();
    }

    private static byte[] BuildSimpleZipBytes() {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "docs/readme.md", "# Stream\n\nFrom non-seekable stream.");
        }

        return ms.ToArray();
    }

    private static byte[] BuildZipWithLargeEntryBytes() {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "big.txt", new string('x', 2048));
        }

        return ms.ToArray();
    }

    private static byte[] BuildZipWithCaseDistinctEntries() {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            const string content = "# Shared\n\nCase-sensitive entry content.";
            WriteTextEntry(archive, "Docs/Readme.md", content);
            WriteTextEntry(archive, "docs/readme.md", content);
        }

        return ms.ToArray();
    }

    private static void WriteTextEntry(ZipArchive archive, string entryPath, string content) {
        var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: false);
        writer.Write(content);
    }

    private static void WriteBytesEntry(ZipArchive archive, string entryPath, byte[] bytes) {
        var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }
}
