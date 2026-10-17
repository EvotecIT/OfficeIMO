using System.Runtime.InteropServices;
using System.Globalization;
using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Zip;

/// <summary>
/// ZIP ingestion adapter for <see cref="OfficeDocumentReader"/>.
/// </summary>
internal static partial class ZipReaderAdapter {
    private static readonly HashSet<string> TextLikeExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml", ".md", ".markdown"
    };

    /// <summary>
    /// Reads supported files from a ZIP archive and emits normalized chunks.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(string zipPath, ReaderOptions? readerOptions = null, ArchiveTraversalOptions? zipOptions = null, CancellationToken cancellationToken = default) {
        return Read(zipPath, readerOptions, zipOptions, readerZipOptions: null, cancellationToken);
    }

    /// <summary>
    /// Reads supported files from a ZIP archive and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(
        string zipPath,
        ReaderOptions? readerOptions,
        ArchiveTraversalOptions? zipOptions,
        ReaderZipOptions? readerZipOptions,
        CancellationToken cancellationToken = default) {
        if (zipPath == null) throw new ArgumentNullException(nameof(zipPath));
        if (zipPath.Length == 0) throw new ArgumentException("ZIP path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"ZIP file '{zipPath}' doesn't exist.", zipPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ArchiveTraversalOptions();
        var effectiveReaderZipOptions = Normalize(readerZipOptions);
        var warningCounter = new WarningCounter();
        ReaderInputLimits.EnforceFileSize(zipPath, effectiveReaderOptions.MaxInputBytes);
        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        long maximumArchiveBytes = Math.Max(1, effectiveZipOptions.MaxArchiveBytes);
        long inputLimit = effectiveReaderOptions.MaxInputBytes.HasValue
            ? Math.Min(effectiveReaderOptions.MaxInputBytes.Value, maximumArchiveBytes)
            : maximumArchiveBytes;
        Stream archiveStream = ReaderInputLimits.EnsureSeekableReadStream(fs, inputLimit,
            cancellationToken, out bool ownsArchiveStream);
        try {
            ArchiveReader.ValidateZipSource(archiveStream, effectiveZipOptions, cancellationToken);
            var archiveSource = BuildArchiveSourceMetadataFromPath(zipPath, archiveStream,
                effectiveReaderOptions.ComputeHashes);
            using var archive = OpenInspection(archiveStream, InspectionOptions(effectiveZipOptions,
                effectiveZipOptions.OperationBudget ?? new ArchiveOperationBudget()), cancellationToken);
            foreach (var chunk in ReadZipArchive(
                         archive,
                         archiveSource,
                         archivePath: archiveSource.Path,
                         readerOptions: effectiveReaderOptions,
                         zipOptions: effectiveZipOptions,
                         readerZipOptions: effectiveReaderZipOptions,
                         warningCounter: warningCounter,
                         cancellationToken: cancellationToken)) {
                yield return chunk;
            }
        } finally {
            if (ownsArchiveStream) archiveStream.Dispose();
        }
    }

    /// <summary>
    /// Reads supported files from a ZIP archive stream and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream zipStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ArchiveTraversalOptions? zipOptions = null,
        CancellationToken cancellationToken = default) {
        return Read(zipStream, sourceName, readerOptions, zipOptions, readerZipOptions: null, cancellationToken);
    }

    /// <summary>
    /// Reads supported files from a ZIP archive stream and emits normalized chunks.
    /// Supports bounded nested ZIP traversal.
    /// </summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream zipStream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ArchiveTraversalOptions? zipOptions,
        ReaderZipOptions? readerZipOptions,
        CancellationToken cancellationToken = default) {
        if (zipStream == null) throw new ArgumentNullException(nameof(zipStream));
        if (!zipStream.CanRead) throw new ArgumentException("ZIP stream must be readable.", nameof(zipStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ArchiveTraversalOptions();
        var effectiveReaderZipOptions = Normalize(readerZipOptions);
        var warningCounter = new WarningCounter();
        var logicalSourceName = "archive.zip";
        if (sourceName != null) {
            var trimmedSourceName = sourceName.Trim();
            if (trimmedSourceName.Length > 0) {
                logicalSourceName = trimmedSourceName;
            }
        }

        long maximumArchiveBytes = Math.Max(1, effectiveZipOptions.MaxArchiveBytes);
        long inputLimit = effectiveReaderOptions.MaxInputBytes.HasValue
            ? Math.Min(effectiveReaderOptions.MaxInputBytes.Value, maximumArchiveBytes)
            : maximumArchiveBytes;
        var archiveStream = ReaderInputLimits.EnsureSeekableReadStream(zipStream, inputLimit, cancellationToken, out var ownsArchiveStream);
        try {
            ArchiveReader.ValidateZipSource(archiveStream, effectiveZipOptions, cancellationToken);
            var archiveSource = BuildArchiveSourceMetadataFromStream(archiveStream, logicalSourceName, effectiveReaderOptions.ComputeHashes);
            using var archive = OpenInspection(archiveStream, InspectionOptions(effectiveZipOptions,
                effectiveZipOptions.OperationBudget ?? new ArchiveOperationBudget()), cancellationToken);
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
        ArchiveInspection archive,
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        ReaderOptions readerOptions,
        ArchiveTraversalOptions zipOptions,
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
        ArchiveInspection archive,
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        int nestedDepth,
        ReaderOptions readerOptions,
        ArchiveTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();

        var traversal = ArchiveReader.TraverseZip(archive, zipOptions, cancellationToken);
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

        var duplicatePaths = traversal.Entries.GroupBy(item => item.FullName, StringComparer.Ordinal)
            .Where(group => group.Count() > 1).Select(group => group.Key);
        var duplicatePathSet = new HashSet<string>(duplicatePaths, StringComparer.Ordinal);
        foreach (var descriptor in traversal.Entries) {
            cancellationToken.ThrowIfCancellationRequested();
            if (descriptor.IsDirectory) continue;

            var entryName = descriptor.FullName;
            var entry = archive.Entries[descriptor.EntryIndex];
            // A second container boundary keeps duplicate identities distinct without changing extensions.
            string entryArchivePath = duplicatePathSet.Contains(entryName)
                ? archivePath + "!/entry-" + descriptor.EntryIndex.ToString("D6", CultureInfo.InvariantCulture)
                : archivePath;
            if (IsZipEntry(entryName)) {
                foreach (var nestedChunk in ReadNestedZipEntry(
                             archive, entry,
                             archiveSource,
                             entryArchivePath,
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
                    entryArchivePath,
                    entryName,
                    warningCounter.Next(),
                    $"Skipped ZIP entry because it exceeds MaxInputBytes ({descriptor.UncompressedLength} > {readerOptions.MaxInputBytes.Value}).",
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: (entry.LastWriteTime ?? DateTimeOffset.MinValue));
                continue;
            }

            if (descriptor.UncompressedLength > int.MaxValue) {
                yield return BuildWarningChunk(
                    archiveSource,
                    entryArchivePath,
                    entryName,
                    warningCounter.Next(),
                    "Skipped ZIP entry because it is too large to materialize in memory.",
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: (entry.LastWriteTime ?? DateTimeOffset.MinValue));
                continue;
            }

            byte[]? bytes = null;
            string? readError = null;
            try {
                bytes = ReadAllBytes(archive, entry, cancellationToken);
            } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
                readError = $"Skipped ZIP entry due read error: {ex.GetType().Name}.";
            }

            if (readError != null) {
                yield return BuildWarningChunk(
                    archiveSource,
                    entryArchivePath,
                    entryName,
                    warningCounter.Next(),
                    readError,
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: (entry.LastWriteTime ?? DateTimeOffset.MinValue));
                continue;
            }

            IEnumerable<ReaderChunk>? chunks = null;
            OfficeDocumentReadResult? nestedDocument = null;
            string? parseError = null;
            try {
                using var content = new MemoryStream(bytes!, 0, bytes!.Length, writable: false, publiclyVisible: true);
                nestedDocument = ReaderNestedContent.ReadDocumentInContainer(content, entryName, BuildVirtualPath(entryArchivePath, entryName), readerOptions, cancellationToken);
                chunks = nestedDocument.Chunks;
            } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
                parseError = $"Skipped ZIP entry due parse error: {ex.GetType().Name}.";
            }

            if (parseError != null) {
                yield return BuildWarningChunk(
                    archiveSource,
                    entryArchivePath,
                    entryName,
                    warningCounter.Next(),
                    parseError,
                    readerOptions.ComputeHashes,
                    sourceLengthBytes: descriptor.UncompressedLength,
                    sourceLastWriteTime: (entry.LastWriteTime ?? DateTimeOffset.MinValue));
                continue;
            }

            var virtualPath = BuildVirtualPath(entryArchivePath, entryName);
            DocumentReaderEngine.ApplyExternalSourceMetadata(nestedDocument!, BuildSourceId(virtualPath),
                NormalizeLastWriteUtc((entry.LastWriteTime ?? DateTimeOffset.MinValue)), descriptor.UncompressedLength, readerOptions.ComputeHashes);
            foreach (var child in chunks!) {
                var chunk = child.CopyForContainer();
                cancellationToken.ThrowIfCancellationRequested();
                ApplyVirtualSourceMetadata(
                    chunk,
                    archiveSource,
                    virtualPath,
                    descriptor.UncompressedLength,
                    (entry.LastWriteTime ?? DateTimeOffset.MinValue),
                    readerOptions.ComputeHashes);
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadNestedZipEntry(
        ArchiveInspection archive, ArchiveInspectionEntry entry,
        ArchiveSourceMetadata archiveSource,
        string archivePath,
        string entryName,
        int nestedDepth,
        ReaderOptions readerOptions,
        ArchiveTraversalOptions zipOptions,
        ReaderZipOptions readerZipOptions,
        WarningCounter warningCounter,
        CancellationToken cancellationToken) {
        var nestedEntryLength = (long?)entry.Length;
        var nestedEntryLastWriteTime = (entry.LastWriteTime ?? DateTimeOffset.MinValue);

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

        using var nestedScope = ReaderReadScope.EnterContainer(readerOptions);
        if (!TryReadAllBytes(archive, entry, cancellationToken, out var nestedBytes, out var readError)) {
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
            ArchiveReader.ValidateZipSource(nestedStream, zipOptions, cancellationToken);
            using var nestedArchive = OpenInspection(nestedStream, InspectionOptions(zipOptions,
                archive.OperationBudget.CreateNested()), cancellationToken);
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
        } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
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
        if (ReaderNestedContent.CanRead(entryName)) return true;
        var kind = DocumentReaderEngine.DetectKind(entryName);
        if (kind != ReaderInputKind.Unknown) return true;

        var ext = Path.GetExtension(entryName);
        if (string.IsNullOrWhiteSpace(ext)) return false;
        return TextLikeExtensions.Contains(ext);
    }

    private static bool IsZipEntry(string entryName) {
        if (string.IsNullOrWhiteSpace(entryName)) return false;
        return string.Equals(Path.GetExtension(entryName), ".zip", StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] ReadAllBytes(ArchiveInspection archive, ArchiveInspectionEntry entry, CancellationToken cancellationToken) {
        ReaderNestedContent.ReserveDecodedInput(entry.Length);
        try {
            using var source = archive.OpenRead(entry.Ordinal, cancellationToken);
            using var output = new MemoryStream(checked((int)entry.Length));
            source.CopyTo(output);
            byte[] payload = output.GetBuffer();
            ReaderNestedContent.MarkReservedPayload(payload);
            return payload;
        } catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }
    }

    private static ArchiveInspection OpenInspection(Stream source, ArchiveInspectionOptions options, CancellationToken token) {
        // Every caller holds either Reader's private snapshot or a verified nested member's immutable
        // bytes until inspection is disposed. Borrow that source instead of making a second snapshot.
        try { return ArchiveInspector.OpenZipSnapshot(source, options, token, leaveOpen: true); }
        catch (OfficeIMO.OfficeInspectionResourceLimitException error) {
            throw new ReaderResourceLimitException(error.LimitName, error.Maximum, error);
        }
    }

    private static ArchiveInspectionOptions InspectionOptions(ArchiveTraversalOptions options, ArchiveOperationBudget budget) => new() {
        OperationBudget = budget, ZipLegacyCodePage = options.ZipLegacyCodePage,
        MaxInputBytes = Math.Max(1, options.MaxArchiveBytes), MaxEntryCount = Math.Max(1, options.MaxPhysicalEntries),
        // Inventory preserves oversized entries so ingestion policy can skip them with a warning.
        MaxEntryBytes = long.MaxValue, MaxExpandedBytes = long.MaxValue
    };

    private static bool TryReadAllBytes(ArchiveInspection archive, ArchiveInspectionEntry entry, CancellationToken cancellationToken, out byte[]? bytes, out string? error) {
        try {
            bytes = ReadAllBytes(archive, entry, cancellationToken);
            error = null;
            return true;
        } catch (Exception ex) when (ex is not OperationCanceledException and not ReaderResourceLimitException and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException) {
            bytes = null;
            error = $"Skipped nested ZIP entry due read error: {ex.GetType().Name}.";
            return false;
        }
    }

}
