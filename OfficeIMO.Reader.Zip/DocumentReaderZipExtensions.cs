using OfficeIMO.Zip;

namespace OfficeIMO.Reader.Zip;

/// <summary>
/// ZIP ingestion adapter for <see cref="DocumentReader"/>.
/// </summary>
public static class DocumentReaderZipExtensions {
    private static readonly HashSet<string> TextLikeExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml", ".md", ".markdown"
    };

    /// <summary>
    /// Reads supported files from a ZIP archive and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadZip(string zipPath, ReaderOptions? readerOptions = null, ZipTraversalOptions? zipOptions = null, CancellationToken cancellationToken = default) {
        return ReadZip(zipPath, readerOptions, zipOptions, readerZipOptions: null, cancellationToken);
    }

    /// <summary>
    /// Reads supported files from a ZIP archive and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadZip(
        string zipPath,
        ReaderOptions? readerOptions,
        ZipTraversalOptions? zipOptions,
        ReaderZipOptions? readerZipOptions,
        CancellationToken cancellationToken = default) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ZipTraversalOptions();
        var effectiveReaderZipOptions = Normalize(readerZipOptions);
        var warningCounter = new WarningCounter();

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        foreach (var chunk in ReadZipArchive(
                     archive,
                     archivePath: zipPath,
                     readerOptions: effectiveReaderOptions,
                     zipOptions: effectiveZipOptions,
                     readerZipOptions: effectiveReaderZipOptions,
                     warningCounter: warningCounter,
                     cancellationToken: cancellationToken)) {
            yield return chunk;
        }
    }

    /// <summary>
    /// Reads supported files from a ZIP archive stream and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadZip(
        Stream zipStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ZipTraversalOptions? zipOptions = null,
        CancellationToken cancellationToken = default) {
        return ReadZip(zipStream, sourceName, readerOptions, zipOptions, readerZipOptions: null, cancellationToken);
    }

    /// <summary>
    /// Reads supported files from a ZIP archive stream and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> ReadZip(
        Stream zipStream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ZipTraversalOptions? zipOptions,
        ReaderZipOptions? readerZipOptions,
        CancellationToken cancellationToken = default) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead) throw new ArgumentException("ZIP stream must be readable.", nameof(zipStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ZipTraversalOptions();
        var effectiveReaderZipOptions = Normalize(readerZipOptions);
        var warningCounter = new WarningCounter();
        var logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "archive.zip" : sourceName!;

        var archiveStream = EnsureSeekableReadStream(zipStream, cancellationToken, out var ownsArchiveStream);
        try {
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (var chunk in ReadZipArchive(
                         archive,
                         archivePath: logicalSourceName,
                         readerOptions: effectiveReaderOptions,
                         zipOptions: effectiveZipOptions,
                         readerZipOptions: effectiveReaderZipOptions,
                         warningCounter: warningCounter,
                         cancellationToken: cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsArchiveStream) {
                archiveStream.Dispose();
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadZipArchive(
        ZipArchive archive,
        string archivePath,
        ReaderOptions readerOptions,
        ZipTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        foreach (var chunk in ReadArchive(
                     archive,
                     archivePath: archivePath,
                     nestedDepth: 0,
                     readerOptions: readerOptions,
                     zipOptions: zipOptions,
                     readerZipOptions: readerZipOptions,
                     warningCounter: warningCounter,
                     cancellationToken: cancellationToken)) {
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadArchive(
        ZipArchive archive,
        string archivePath,
        int nestedDepth,
        ReaderOptions readerOptions,
        ZipTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        var traversal = ZipTraversal.Traverse(archive, zipOptions);
        foreach (var traversalWarning in traversal.Warnings) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return BuildWarningChunk(
                archivePath,
                traversalWarning.EntryPath,
                warningCounter.Next(),
                traversalWarning.Warning);
        }

        foreach (var descriptor in traversal.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (descriptor.IsDirectory) continue;

            var entryName = descriptor.FullName;
            var entry = archive.GetEntry(entryName);
            if (entry == null) {
                yield return BuildWarningChunk(
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    "Skipped ZIP entry because it could not be opened from archive index.");
                continue;
            }

            if (IsZipEntry(entryName)) {
                foreach (var nestedChunk in ReadNestedZipEntry(
                             entry,
                             archivePath,
                             entryName,
                             nestedDepth,
                             readerOptions,
                             zipOptions,
                             readerZipOptions,
                             warningCounter,
                             cancellationToken)) {
                    yield return nestedChunk;
                }

                continue;
            }

            if (!ShouldAttemptRead(entryName)) continue;

            if (readerOptions.MaxInputBytes.HasValue && descriptor.UncompressedLength > readerOptions.MaxInputBytes.Value) {
                yield return BuildWarningChunk(
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    $"Skipped ZIP entry because it exceeds MaxInputBytes ({descriptor.UncompressedLength} > {readerOptions.MaxInputBytes.Value}).");
                continue;
            }

            if (descriptor.UncompressedLength > int.MaxValue) {
                yield return BuildWarningChunk(
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    "Skipped ZIP entry because it is too large to materialize in memory.");
                continue;
            }

            byte[]? bytes = null;
            string? readError = null;
            try {
                bytes = ReadAllBytes(entry, cancellationToken);
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                readError = $"Skipped ZIP entry due read error: {ex.GetType().Name}.";
            }

            if (readError != null) {
                yield return BuildWarningChunk(archivePath, entryName, warningCounter.Next(), readError);
                continue;
            }

            IEnumerable<ReaderChunk>? chunks = null;
            string? parseError = null;
            try {
                chunks = DocumentReader.Read(bytes!, entryName, readerOptions, cancellationToken);
            } catch (Exception ex) when (ex is not OperationCanceledException) {
                parseError = $"Skipped ZIP entry due parse error: {ex.GetType().Name}.";
            }

            if (parseError != null) {
                yield return BuildWarningChunk(archivePath, entryName, warningCounter.Next(), parseError);
                continue;
            }

            var virtualPath = BuildVirtualPath(archivePath, entryName);
            foreach (var chunk in chunks!) {
                cancellationToken.ThrowIfCancellationRequested();
                chunk.Location.Path = virtualPath;
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadNestedZipEntry(
        ZipArchiveEntry entry,
        string archivePath,
        string entryName,
        int nestedDepth,
        ReaderOptions readerOptions,
        ZipTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        if (!readerZipOptions.ReadNestedZipEntries) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                "Skipped nested ZIP entry because ReadNestedZipEntries is disabled.");
            yield break;
        }

        if (nestedDepth >= readerZipOptions.MaxNestedDepth) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because MaxNestedDepth ({readerZipOptions.MaxNestedDepth}) was reached.");
            yield break;
        }

        if (!TryGetEntryLength(entry, out var entryLength)) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                "Skipped nested ZIP entry because size metadata could not be read.");
            yield break;
        }

        if (readerZipOptions.MaxNestedArchiveBytes.HasValue && entryLength > readerZipOptions.MaxNestedArchiveBytes.Value) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because size {entryLength} exceeds MaxNestedArchiveBytes ({readerZipOptions.MaxNestedArchiveBytes.Value}).");
            yield break;
        }

        if (readerOptions.MaxInputBytes.HasValue && entryLength > readerOptions.MaxInputBytes.Value) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because it exceeds MaxInputBytes ({entryLength} > {readerOptions.MaxInputBytes.Value}).");
            yield break;
        }

        if (!TryReadAllBytes(entry, cancellationToken, out var nestedBytes, out var readError)) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                readError ?? "Skipped nested ZIP entry due read error.");
            yield break;
        }

        List<ReaderChunk>? nestedChunks = null;
        string? parseError = null;
        try {
            using var nestedStream = new MemoryStream(nestedBytes!, writable: false);
            using var nestedArchive = new ZipArchive(nestedStream, ZipArchiveMode.Read, leaveOpen: false);
            var nestedArchivePath = BuildVirtualPath(archivePath, entryName);

            nestedChunks = ReadArchive(
                         nestedArchive,
                         nestedArchivePath,
                         nestedDepth + 1,
                         readerOptions,
                         zipOptions,
                         readerZipOptions,
                         warningCounter,
                         cancellationToken).ToList();
        } catch (Exception ex) when (ex is not OperationCanceledException) {
            parseError = $"Skipped nested ZIP entry due archive parse error: {ex.GetType().Name}.";
        }

        if (parseError != null) {
            yield return BuildWarningChunk(
                archivePath,
                entryName,
                warningCounter.Next(),
                parseError);
            yield break;
        }

        foreach (var nestedChunk in nestedChunks!) {
            yield return nestedChunk;
        }
    }

    private static ReaderZipOptions Normalize(ReaderZipOptions? options) {
        var source = options ?? new ReaderZipOptions();
        var normalized = new ReaderZipOptions {
            ReadNestedZipEntries = source.ReadNestedZipEntries,
            MaxNestedDepth = source.MaxNestedDepth,
            MaxNestedArchiveBytes = source.MaxNestedArchiveBytes
        };

        if (normalized.MaxNestedDepth < 0) normalized.MaxNestedDepth = 0;
        if (normalized.MaxNestedArchiveBytes.HasValue && normalized.MaxNestedArchiveBytes.Value < 1) {
            normalized.MaxNestedArchiveBytes = 1;
        }

        return normalized;
    }

    private static bool ShouldAttemptRead(string entryName) {
        var kind = DocumentReader.DetectKind(entryName);
        if (kind != ReaderInputKind.Unknown) return true;

        var ext = Path.GetExtension(entryName);
        if (string.IsNullOrWhiteSpace(ext)) return false;
        return TextLikeExtensions.Contains(ext);
    }

    private static bool IsZipEntry(string entryName) {
        if (string.IsNullOrWhiteSpace(entryName)) return false;
        return string.Equals(Path.GetExtension(entryName), ".zip", StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] ReadAllBytes(ZipArchiveEntry entry, CancellationToken cancellationToken) {
        using var source = entry.Open();
        using var ms = entry.Length > 0 && entry.Length < int.MaxValue
            ? new MemoryStream((int)entry.Length)
            : new MemoryStream();

        var buffer = new byte[64 * 1024];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            var read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            ms.Write(buffer, 0, read);
        }

        return ms.ToArray();
    }

    private static Stream EnsureSeekableReadStream(Stream stream, CancellationToken cancellationToken, out bool ownsStream) {
        if (stream.CanSeek) {
            ownsStream = false;
            return stream;
        }

        var buffer = new MemoryStream();
        var chunk = new byte[64 * 1024];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            var read = stream.Read(chunk, 0, chunk.Length);
            if (read <= 0) break;
            buffer.Write(chunk, 0, read);
        }

        buffer.Position = 0;
        ownsStream = true;
        return buffer;
    }

    private static bool TryReadAllBytes(ZipArchiveEntry entry, CancellationToken cancellationToken, out byte[]? bytes, out string? error) {
        try {
            bytes = ReadAllBytes(entry, cancellationToken);
            error = null;
            return true;
        } catch (Exception ex) when (ex is not OperationCanceledException) {
            bytes = null;
            error = $"Skipped nested ZIP entry due read error: {ex.GetType().Name}.";
            return false;
        }
    }

    private static bool TryGetEntryLength(ZipArchiveEntry entry, out long length) {
        try {
            length = entry.Length;
            return true;
        } catch {
            length = 0;
            return false;
        }
    }

    private static ReaderChunk BuildWarningChunk(string archivePath, string entryName, int warningIndex, string warning) {
        var warningPath = string.IsNullOrWhiteSpace(entryName)
            ? archivePath
            : BuildVirtualPath(archivePath, entryName);

        return new ReaderChunk {
            Id = $"zip-warning-{warningIndex.ToString("D4", System.Globalization.CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation {
                Path = warningPath,
                BlockIndex = warningIndex
            },
            Text = warning,
            Warnings = new[] { warning }
        };
    }

    private static string BuildVirtualPath(string zipPath, string entryName) {
        return zipPath + "::" + entryName.Replace('\\', '/');
    }

    private sealed class WarningCounter {
        public int Value { get; set; }
        public int Next() {
            var current = Value;
            Value = current + 1;
            return current;
        }
    }
}
