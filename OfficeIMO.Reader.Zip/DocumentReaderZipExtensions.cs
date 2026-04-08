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
        ReaderInputLimits.EnforceFileSize(zipPath, effectiveReaderOptions.MaxInputBytes);
        var archiveSource = BuildArchiveSourceMetadataFromPath(zipPath, effectiveReaderOptions.ComputeHashes);

        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var archive = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        foreach (var chunk in ReadZipArchive(
                     archive,
                     archiveSource,
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
        var logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "archive.zip" : sourceName.Trim();

        var archiveStream = ReaderInputLimits.EnsureSeekableReadStream(zipStream, effectiveReaderOptions.MaxInputBytes, cancellationToken, out var ownsArchiveStream);
        try {
            var archiveSource = BuildArchiveSourceMetadataFromStream(archiveStream, logicalSourceName, effectiveReaderOptions.ComputeHashes);
            using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read, leaveOpen: true);
            foreach (var chunk in ReadZipArchive(
                         archive,
                         archiveSource,
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
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        ReaderOptions readerOptions,
        ZipTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        foreach (var chunk in ReadArchive(
                     archive,
                     archiveSource,
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
        ArchiveSourceMetadata archiveSource,
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
                archiveSource,
                archivePath,
                traversalWarning.EntryPath,
                warningCounter.Next(),
                traversalWarning.Warning,
                readerOptions.ComputeHashes);
        }

        foreach (var descriptor in traversal.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (descriptor.IsDirectory) continue;

            var entryName = descriptor.FullName;
            var entry = archive.GetEntry(entryName);
            if (entry == null) {
                yield return BuildWarningChunk(
                    archiveSource,
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    "Skipped ZIP entry because it could not be opened from archive index.",
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength);
                continue;
            }

            if (IsZipEntry(entryName)) {
                foreach (var nestedChunk in ReadNestedZipEntry(
                             entry,
                             archiveSource,
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
                    archiveSource,
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    $"Skipped ZIP entry because it exceeds MaxInputBytes ({descriptor.UncompressedLength} > {readerOptions.MaxInputBytes.Value}).",
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: entry.LastWriteTime);
                continue;
            }

            if (descriptor.UncompressedLength > int.MaxValue) {
                yield return BuildWarningChunk(
                    archiveSource,
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    "Skipped ZIP entry because it is too large to materialize in memory.",
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: entry.LastWriteTime);
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
                yield return BuildWarningChunk(
                    archiveSource,
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    readError,
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: entry.LastWriteTime);
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
                yield return BuildWarningChunk(
                    archiveSource,
                    archivePath,
                    entryName,
                    warningCounter.Next(),
                    parseError,
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: entry.LastWriteTime);
                continue;
            }

            var virtualPath = BuildVirtualPath(archivePath, entryName);
            foreach (var chunk in chunks!) {
                cancellationToken.ThrowIfCancellationRequested();
                ApplyVirtualSourceMetadata(
                    chunk,
                    archiveSource,
                    virtualPath,
                    descriptor.UncompressedLength,
                    entry.LastWriteTime,
                    readerOptions.ComputeHashes);
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadNestedZipEntry(
        ZipArchiveEntry entry,
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        string entryName,
        int nestedDepth,
        ReaderOptions readerOptions,
        ZipTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        var nestedEntryLength = TryGetEntryLength(entry, out var resolvedEntryLength) ? resolvedEntryLength : (long?)null;
        var nestedEntryLastWriteTime = entry.LastWriteTime;

        if (!readerZipOptions.ReadNestedZipEntries) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                "Skipped nested ZIP entry because ReadNestedZipEntries is disabled.",
                readerOptions.ComputeHashes,
                sourceLengthBytes: nestedEntryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
            yield break;
        }

        if (nestedDepth >= readerZipOptions.MaxNestedDepth) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because MaxNestedDepth ({readerZipOptions.MaxNestedDepth}) was reached.",
                readerOptions.ComputeHashes,
                sourceLengthBytes: nestedEntryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
            yield break;
        }

        if (!nestedEntryLength.HasValue) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                "Skipped nested ZIP entry because size metadata could not be read.",
                readerOptions.ComputeHashes,
                sourceLastWriteTime: nestedEntryLastWriteTime);
            yield break;
        }

        var entryLength = nestedEntryLength.Value;

        if (readerZipOptions.MaxNestedArchiveBytes.HasValue && entryLength > readerZipOptions.MaxNestedArchiveBytes.Value) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because size {entryLength} exceeds MaxNestedArchiveBytes ({readerZipOptions.MaxNestedArchiveBytes.Value}).",
                readerOptions.ComputeHashes,
                sourceLengthBytes: entryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
            yield break;
        }

        if (readerOptions.MaxInputBytes.HasValue && entryLength > readerOptions.MaxInputBytes.Value) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                $"Skipped nested ZIP entry because it exceeds MaxInputBytes ({entryLength} > {readerOptions.MaxInputBytes.Value}).",
                readerOptions.ComputeHashes,
                sourceLengthBytes: entryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
            yield break;
        }

        if (!TryReadAllBytes(entry, cancellationToken, out var nestedBytes, out var readError)) {
            yield return BuildWarningChunk(
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                readError ?? "Skipped nested ZIP entry due read error.",
                readerOptions.ComputeHashes,
                sourceLengthBytes: entryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
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
                         archiveSource,
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
                archiveSource,
                archivePath,
                entryName,
                warningCounter.Next(),
                parseError,
                readerOptions.ComputeHashes,
                sourceLengthBytes: entryLength,
                sourceLastWriteTime: nestedEntryLastWriteTime);
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

    private static ReaderChunk BuildWarningChunk(
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        string entryName,
        int warningIndex,
        string warning,
        bool computeHashes,
        long? sourceLengthBytes = null,
        DateTimeOffset? sourceLastWriteTime = null) {
        var warningPath = string.IsNullOrWhiteSpace(entryName)
            ? archivePath
            : BuildVirtualPath(archivePath, entryName);

        var chunk = new ReaderChunk {
            Id = $"zip-warning-{warningIndex.ToString("D4", System.Globalization.CultureInfo.InvariantCulture)}",
            Kind = ReaderInputKind.Zip,
            Location = new ReaderLocation {
                Path = warningPath,
                BlockIndex = warningIndex
            },
            Text = warning,
            Warnings = new[] { warning }
        };

        ApplyWarningSourceMetadata(chunk, archiveSource, warningPath, sourceLengthBytes, sourceLastWriteTime, computeHashes);
        return chunk;
    }

    private static string BuildVirtualPath(string zipPath, string entryName) {
        return zipPath + "::" + entryName.Replace('\\', '/');
    }

    private static void ApplyVirtualSourceMetadata(
        ReaderChunk chunk,
        ArchiveSourceMetadata archiveSource,
        string virtualPath,
        long uncompressedLength,
        DateTimeOffset lastWriteTime,
        bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));

        chunk.Location.Path = virtualPath;
        chunk.SourceId = BuildSourceId(virtualPath);
        chunk.SourceHash ??= archiveSource.SourceHash;
        chunk.SourceLengthBytes = uncompressedLength >= 0 ? uncompressedLength : chunk.SourceLengthBytes;

        var sourceLastWriteUtc = NormalizeLastWriteUtc(lastWriteTime);
        if (sourceLastWriteUtc.HasValue) {
            chunk.SourceLastWriteUtc = sourceLastWriteUtc;
        }
        chunk.SourceLastWriteUtc ??= archiveSource.LastWriteUtc;

        if (!chunk.TokenEstimate.HasValue) {
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        }

        if (computeHashes) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }
    }

    private static void ApplyWarningSourceMetadata(
        ReaderChunk chunk,
        ArchiveSourceMetadata archiveSource,
        string sourcePath,
        long? sourceLengthBytes,
        DateTimeOffset? sourceLastWriteTime,
        bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));

        chunk.Location.Path = sourcePath;
        chunk.SourceId = BuildSourceId(sourcePath);
        chunk.SourceHash ??= archiveSource.SourceHash;

        if (sourceLengthBytes.HasValue && sourceLengthBytes.Value >= 0) {
            chunk.SourceLengthBytes = sourceLengthBytes.Value;
        } else if (string.Equals(sourcePath, archiveSource.Path, StringComparison.OrdinalIgnoreCase)) {
            chunk.SourceLengthBytes ??= archiveSource.LengthBytes;
        }

        if (sourceLastWriteTime.HasValue) {
            var normalized = NormalizeLastWriteUtc(sourceLastWriteTime.Value);
            if (normalized.HasValue) {
                chunk.SourceLastWriteUtc = normalized.Value;
            }
        } else if (string.Equals(sourcePath, archiveSource.Path, StringComparison.OrdinalIgnoreCase)) {
            chunk.SourceLastWriteUtc ??= archiveSource.LastWriteUtc;
        }

        if (!chunk.TokenEstimate.HasValue) {
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        }

        if (computeHashes) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }
    }

    private static int EstimateTokenCount(string? text) {
        var safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static ArchiveSourceMetadata BuildArchiveSourceMetadataFromPath(string zipPath, bool computeHash) {
        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fileInfo = new FileInfo(zipPath);
            if (fileInfo.Exists) {
                lastWriteUtc = fileInfo.LastWriteTimeUtc;
                lengthBytes = fileInfo.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        return new ArchiveSourceMetadata {
            Path = zipPath,
            SourceHash = computeHash ? TryComputeFileSha256(zipPath) : null,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static ArchiveSourceMetadata BuildArchiveSourceMetadataFromStream(Stream stream, string sourceName, bool computeHash) {
        long? lengthBytes = null;
        try {
            if (stream.CanSeek) {
                lengthBytes = stream.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        return new ArchiveSourceMetadata {
            Path = sourceName,
            SourceHash = computeHash ? TryComputeStreamSha256(stream) : null,
            LastWriteUtc = null,
            LengthBytes = lengthBytes
        };
    }

    private static DateTime? NormalizeLastWriteUtc(DateTimeOffset lastWriteTime) {
        if (lastWriteTime == default) return null;

        // Zip archives sometimes surface sentinel timestamps; keep only practical values.
        var utc = lastWriteTime.UtcDateTime;
        if (utc.Year < 1980) return null;
        return utc;
    }

    private static string BuildSourceId(string sourceKey) {
        var normalized = NormalizeSourceKeyForId(sourceKey);

        return "src:" + ComputeSha256Hex(normalized);
    }

    private static string NormalizeSourceKeyForId(string? sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (Path.DirectorySeparatorChar != '\\') {
            return normalized;
        }

        var virtualPathSeparatorIndex = normalized.IndexOf("::", StringComparison.Ordinal);
        if (virtualPathSeparatorIndex < 0) {
            return normalized.ToLowerInvariant();
        }

        var archivePath = normalized.Substring(0, virtualPathSeparatorIndex).ToLowerInvariant();
        var archiveEntryPath = normalized.Substring(virtualPathSeparatorIndex);
        return archivePath + archiveEntryPath;
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        var data = string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.HeadingPath ?? string.Empty,
            chunk.Location.HeadingSlug ?? string.Empty,
            chunk.Location.SourceBlockKind ?? string.Empty,
            chunk.Location.BlockAnchor ?? string.Empty,
            chunk.Location.Sheet ?? string.Empty,
            chunk.Location.A1Range ?? string.Empty,
            chunk.Location.Page?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.Slide?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.StartLine?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedStartLine?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.NormalizedEndLine?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty);

        return ComputeSha256Hex(data);
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var bytes = System.Text.Encoding.UTF8.GetBytes(value ?? string.Empty);
        var hash = sha.ComputeHash(bytes);

        var sb = new System.Text.StringBuilder(hash.Length * 2);
        for (int i = 0; i < hash.Length; i++) {
            sb.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static string ComputeSha256Hex(Stream stream) {
        using var sha = System.Security.Cryptography.SHA256.Create();
        var hash = sha.ComputeHash(stream);

        var sb = new System.Text.StringBuilder(hash.Length * 2);
        for (int i = 0; i < hash.Length; i++) {
            sb.Append(hash[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }

    private static string? TryComputeFileSha256(string path) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs);
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream) {
        if (stream == null || !stream.CanSeek) return null;

        long position;
        try {
            position = stream.Position;
        } catch {
            return null;
        }

        try {
            stream.Position = 0;
            var hash = ComputeSha256Hex(stream);
            stream.Position = position;
            return hash;
        } catch {
            try {
                stream.Position = position;
            } catch {
                // ignore
            }

            return null;
        }
    }

    private sealed class WarningCounter {
        public int Value { get; set; }
        public int Next() {
            var current = Value;
            Value = current + 1;
            return current;
        }
    }

    private sealed class ArchiveSourceMetadata {
        public string Path { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
