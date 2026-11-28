using System.Globalization;
using OfficeIMO.Archives;

namespace OfficeIMO.Reader.Archives;

/// <summary>
/// Archive ingestion adapter for <see cref="OfficeDocumentReader"/>: ZIP, 7z, Microsoft Cabinet, TAR plain or compressed,
/// single compressed files and, when enabled, Windows Installer packages. Nested archives of any of these kinds are read
/// through the same path and limits.
/// </summary>
internal static partial class ArchiveReaderAdapter {
    private static readonly HashSet<string> TextLikeExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml", ".md", ".markdown"
    };

    private static readonly HashSet<string> ArchiveExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        ".zip", ".7z", ".cab", ".tar", ".tgz", ".gz", ".tzst", ".zst", ".txz", ".xz", ".tlz", ".lz"
    };

    /// <summary>Reads supported files from an archive and emits normalized chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(string zipPath, ReaderOptions? readerOptions = null, ArchiveTraversalOptions? zipOptions = null, CancellationToken cancellationToken = default) {
        return Read(zipPath, readerOptions, zipOptions, readerArchiveOptions: null, cancellationToken);
    }

    /// <summary>Reads supported files from an archive and emits normalized chunks, with bounded nested archive traversal.</summary>
    public static IEnumerable<ReaderChunk> Read(
        string zipPath,
        ReaderOptions? readerOptions,
        ArchiveTraversalOptions? zipOptions,
        ReaderArchiveOptions? readerArchiveOptions,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(zipPath);
        if (zipPath.Length == 0) throw new ArgumentException("Archive path cannot be empty.", nameof(zipPath));
        if (!File.Exists(zipPath)) throw new FileNotFoundException($"Archive file '{zipPath}' doesn't exist.", zipPath);

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ArchiveTraversalOptions();
        ReaderInputLimits.EnforceFileSize(zipPath, effectiveReaderOptions.MaxInputBytes);
        using var fs = new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        Stream archiveStream = ReaderInputLimits.EnsureSeekableReadStream(fs, InputLimit(effectiveReaderOptions, effectiveZipOptions),
            cancellationToken, out bool ownsArchiveStream);
        try {
            var source = BuildArchiveSourceMetadataFromPath(zipPath, archiveStream, effectiveReaderOptions.ComputeHashes);
            var context = new ReadContext(source, effectiveReaderOptions, effectiveZipOptions, Normalize(readerArchiveOptions), cancellationToken);
            foreach (var chunk in ReadContainer(archiveStream, source.Path, nestedDepth: 0,
                         effectiveZipOptions.OperationBudget ?? new ArchiveOperationBudget(), context)) {
                yield return chunk;
            }
        } finally {
            if (ownsArchiveStream) archiveStream.Dispose();
        }
    }

    /// <summary>Reads supported files from an archive stream and emits normalized chunks.</summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream zipStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ArchiveTraversalOptions? zipOptions = null,
        CancellationToken cancellationToken = default) {
        return Read(zipStream, sourceName, readerOptions, zipOptions, readerArchiveOptions: null, cancellationToken);
    }

    /// <summary>Reads supported files from an archive stream and emits normalized chunks, with bounded nested archive traversal.</summary>
    public static IEnumerable<ReaderChunk> Read(
        Stream zipStream,
        string? sourceName,
        ReaderOptions? readerOptions,
        ArchiveTraversalOptions? zipOptions,
        ReaderArchiveOptions? readerArchiveOptions,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(zipStream);
        if (!zipStream.CanRead) throw new ArgumentException("Archive stream must be readable.", nameof(zipStream));

        var effectiveReaderOptions = readerOptions ?? new ReaderOptions();
        var effectiveZipOptions = zipOptions ?? new ArchiveTraversalOptions();
        string logicalSourceName = string.IsNullOrWhiteSpace(sourceName) ? "archive.zip" : sourceName!.Trim();
        var archiveStream = ReaderInputLimits.EnsureSeekableReadStream(zipStream, InputLimit(effectiveReaderOptions, effectiveZipOptions),
            cancellationToken, out var ownsArchiveStream);
        try {
            var source = BuildArchiveSourceMetadataFromStream(archiveStream, logicalSourceName, effectiveReaderOptions.ComputeHashes);
            var context = new ReadContext(source, effectiveReaderOptions, effectiveZipOptions, Normalize(readerArchiveOptions), cancellationToken);
            foreach (var chunk in ReadContainer(archiveStream, logicalSourceName, nestedDepth: 0,
                         effectiveZipOptions.OperationBudget ?? new ArchiveOperationBudget(), context)) {
                yield return chunk;
            }
        } finally {
            if (ownsArchiveStream) archiveStream.Dispose();
        }
    }

    private static long InputLimit(ReaderOptions readerOptions, ArchiveTraversalOptions zipOptions) {
        long maximumArchiveBytes = Math.Max(1, zipOptions.MaxArchiveBytes);
        return readerOptions.MaxInputBytes.HasValue ? Math.Min(readerOptions.MaxInputBytes.Value, maximumArchiveBytes) : maximumArchiveBytes;
    }

    /// <summary>Reads one container from a seekable snapshot, choosing the format from its content.</summary>
    private static IEnumerable<ReaderChunk> ReadContainer(Stream source, string archivePath, int nestedDepth, ArchiveOperationBudget budget,
        ReadContext context) {
        context.Token.ThrowIfCancellationRequested();
        ArchiveFormat? format = DetectFormat(source, context.Token);
        switch (format) {
            case ArchiveFormat.Tar:
            case ArchiveFormat.TarGZip:
            case ArchiveFormat.TarZstandard:
            case ArchiveFormat.TarXz:
            case ArchiveFormat.TarLzip:
            case ArchiveFormat.GZip:
            case ArchiveFormat.Zstandard:
            case ArchiveFormat.Xz:
            case ArchiveFormat.Lzip:
                return ReadTarOrCompressedFile(source, format.Value, archivePath, nestedDepth, budget, context);
            case ArchiveFormat.SevenZip:
                return ReadSevenZip(source, archivePath, nestedDepth, budget, context);
            case ArchiveFormat.Cab:
                return ReadInspected(source, archivePath, nestedDepth, budget, context);
            case ArchiveFormat.Msi:
                return context.Archive.ReadInstallerPackages
                    ? ReadInspected(source, archivePath, nestedDepth, budget, context)
                    : new[] {
                        BuildWarningChunk(context.Source, archivePath, string.Empty, context.Warnings.Next(),
                            "Skipped Windows Installer package because ReadInstallerPackages is disabled.", context.ReaderOptions.ComputeHashes,
                            sourceLengthBytes: source.Length)
                    };
            default:
                // ZIP, and anything unrecognized: ZIP validation reports why the input is not a readable archive.
                return ReadZip(source, archivePath, nestedDepth, budget, context);
        }
    }

    private static ArchiveFormat? DetectFormat(Stream source, CancellationToken token) {
        try {
            return ArchiveExtractor.Detect(source, token);
        } catch (NotSupportedException) {
            return null;
        }
    }

    private static IEnumerable<ReaderChunk> ReadZip(Stream source, string archivePath, int nestedDepth, ArchiveOperationBudget budget,
        ReadContext context) {
        ArchiveReader.ValidateZipSource(source, context.Traversal, context.Token);
        using var archive = OpenInspection(source, InspectionOptions(context.Traversal, budget), context.Token);
        var traversal = ArchiveReader.TraverseZip(archive, context.Traversal, context.Token);
        foreach (var traversalWarning in traversal.Warnings) {
            context.Token.ThrowIfCancellationRequested();
            yield return BuildWarningChunk(context.Source, archivePath, traversalWarning.EntryPath, context.Warnings.Next(),
                traversalWarning.Warning, context.ReaderOptions.ComputeHashes);
        }

        var duplicatePathSet = new HashSet<string>(traversal.Entries.GroupBy(item => item.FullName, StringComparer.Ordinal)
            .Where(group => group.Count() > 1).Select(group => group.Key), StringComparer.Ordinal);
        foreach (var descriptor in traversal.Entries) {
            context.Token.ThrowIfCancellationRequested();
            if (descriptor.IsDirectory) continue;
            var entry = archive.Entries[descriptor.EntryIndex];
            // Traversal already applies ZipArchive's local-zone interpretation of DOS wall time.
            var lastWriteTime = new DateTimeOffset(descriptor.LastWriteUtc, TimeSpan.Zero);
            string entryArchivePath = duplicatePathSet.Contains(descriptor.FullName)
                ? DuplicatePath(archivePath, descriptor.EntryIndex) : archivePath;
            foreach (var chunk in EmitMember(descriptor.FullName, entryArchivePath, descriptor.UncompressedLength, lastWriteTime,
                         () => ReadAllBytes(archive, entry, context.Token), nestedDepth, archive.OperationBudget, context)) {
                yield return chunk;
            }
        }
    }

    /// <summary>A second container boundary keeps duplicate identities distinct without changing extensions.</summary>
    private static string DuplicatePath(string archivePath, int index) =>
        archivePath + "!/entry-" + index.ToString("D6", CultureInfo.InvariantCulture);

    /// <summary>Emits one archive member: a nested archive is read recursively, a supported document is parsed, and
    /// anything that cannot be read becomes a warning chunk. <paramref name="read"/> reserves Reader capacity before it
    /// allocates the payload.</summary>
    private static IEnumerable<ReaderChunk> EmitMember(string entryName, string archivePath, long length, DateTimeOffset? lastWriteTime,
        Func<byte[]> read, int nestedDepth, ArchiveOperationBudget budget, ReadContext context) {
        if (IsNestedArchive(entryName, context.Archive)) {
            foreach (var chunk in ReadNested(entryName, archivePath, length, lastWriteTime, read, nestedDepth, budget, context)) yield return chunk;
            yield break;
        }
        if (!ShouldAttemptRead(entryName)) yield break;

        ReaderOptions readerOptions = context.ReaderOptions;
        string? skip = readerOptions.MaxInputBytes.HasValue && length > readerOptions.MaxInputBytes.Value
            ? $"Skipped archive entry because it exceeds MaxInputBytes ({length} > {readerOptions.MaxInputBytes.Value})."
            : length > int.MaxValue ? "Skipped archive entry because it is too large to materialize in memory." : null;
        byte[]? bytes = null;
        if (skip == null) {
            try {
                bytes = read();
            } catch (Exception ex) when (IsRecoverable(ex)) {
                skip = $"Skipped archive entry due read error: {ex.GetType().Name}.";
            }
        }
        if (skip != null) {
            yield return BuildWarningChunk(context.Source, archivePath, entryName, context.Warnings.Next(), skip, readerOptions.ComputeHashes,
                sourceLengthBytes: length, sourceLastWriteTime: lastWriteTime);
            yield break;
        }

        string virtualPath = BuildVirtualPath(archivePath, entryName);
        OfficeDocumentReadResult? document = null;
        string? parseError = null;
        try {
            using var content = new MemoryStream(bytes!, 0, (int)length, writable: false, publiclyVisible: true);
            document = ReaderNestedContent.ReadDocumentInContainer(content, entryName, virtualPath, readerOptions, context.Token);
        } catch (Exception ex) when (IsRecoverable(ex)) {
            parseError = $"Skipped archive entry due parse error: {ex.GetType().Name}.";
        }
        if (parseError != null) {
            yield return BuildWarningChunk(context.Source, archivePath, entryName, context.Warnings.Next(), parseError, readerOptions.ComputeHashes,
                sourceLengthBytes: length, sourceLastWriteTime: lastWriteTime);
            yield break;
        }

        DocumentReaderEngine.ApplyExternalSourceMetadata(document!, BuildSourceId(virtualPath),
            lastWriteTime.HasValue ? NormalizeLastWriteUtc(lastWriteTime.Value) : null, length, readerOptions.ComputeHashes);
        foreach (var child in document!.Chunks) {
            context.Token.ThrowIfCancellationRequested();
            var chunk = child.CopyForContainer();
            ApplyVirtualSourceMetadata(chunk, context.Source, virtualPath, length, lastWriteTime ?? default, readerOptions.ComputeHashes);
            yield return chunk;
        }
    }

    private static IEnumerable<ReaderChunk> ReadNested(string entryName, string archivePath, long length, DateTimeOffset? lastWriteTime,
        Func<byte[]> read, int nestedDepth, ArchiveOperationBudget budget, ReadContext context) {
        ReaderArchiveOptions options = context.Archive;
        ReaderOptions readerOptions = context.ReaderOptions;
        string? skip = !options.ReadNestedArchives ? "Skipped nested archive because ReadNestedArchives is disabled."
            : nestedDepth >= options.MaxNestedDepth ? $"Skipped nested archive because MaxNestedDepth ({options.MaxNestedDepth}) was reached."
            : options.MaxNestedArchiveBytes.HasValue && length > options.MaxNestedArchiveBytes.Value
                ? $"Skipped nested archive because size {length} exceeds MaxNestedArchiveBytes ({options.MaxNestedArchiveBytes.Value})."
            : readerOptions.MaxInputBytes.HasValue && length > readerOptions.MaxInputBytes.Value
                ? $"Skipped nested archive because it exceeds MaxInputBytes ({length} > {readerOptions.MaxInputBytes.Value})."
            : length > int.MaxValue ? "Skipped nested archive because it is too large to materialize in memory." : null;
        if (skip != null) {
            yield return BuildWarningChunk(context.Source, archivePath, entryName, context.Warnings.Next(), skip, readerOptions.ComputeHashes,
                sourceLengthBytes: length, sourceLastWriteTime: lastWriteTime);
            yield break;
        }

        using var nestedScope = ReaderReadScope.EnterContainer(readerOptions);
        List<ReaderChunk>? nestedChunks = null;
        string? error = null;
        byte[]? bytes = null;
        try {
            bytes = read();
        } catch (Exception ex) when (IsRecoverable(ex)) {
            error = $"Skipped nested archive due read error: {ex.GetType().Name}.";
        }
        if (error == null) {
            try {
                using var nestedStream = new MemoryStream(bytes!, 0, (int)length, writable: false);
                nestedChunks = ReadContainer(nestedStream, BuildVirtualPath(archivePath, entryName), nestedDepth + 1, budget.CreateNested(), context).ToList();
            } catch (Exception ex) when (IsRecoverable(ex)) {
                error = $"Skipped nested archive due archive parse error: {ex.GetType().Name}.";
            }
        }
        if (error != null) {
            yield return BuildWarningChunk(context.Source, archivePath, entryName, context.Warnings.Next(), error, readerOptions.ComputeHashes,
                sourceLengthBytes: length, sourceLastWriteTime: lastWriteTime);
            yield break;
        }
        foreach (var chunk in nestedChunks!) yield return chunk;
    }

    private static bool IsRecoverable(Exception error) => error is not OperationCanceledException and not ReaderResourceLimitException
        and not ReaderAdmissionLimitException and not OfficeIMO.OfficeInspectionResourceLimitException;

    private static ReaderArchiveOptions Normalize(ReaderArchiveOptions? options) {
        var source = options ?? new ReaderArchiveOptions();
        return new ReaderArchiveOptions {
            ReadNestedArchives = source.ReadNestedArchives,
            MaxNestedDepth = Math.Max(0, source.MaxNestedDepth),
            MaxNestedArchiveBytes = source.MaxNestedArchiveBytes.HasValue ? Math.Max(1, source.MaxNestedArchiveBytes.Value) : null,
            ReadInstallerPackages = source.ReadInstallerPackages
        };
    }

    private static bool ShouldAttemptRead(string entryName) {
        if (ReaderNestedContent.CanRead(entryName)) return true;
        var kind = DocumentReaderEngine.DetectKind(entryName);
        if (kind != ReaderInputKind.Unknown) return true;

        var ext = Path.GetExtension(entryName);
        if (string.IsNullOrWhiteSpace(ext)) return false;
        return TextLikeExtensions.Contains(ext);
    }

    private static bool IsNestedArchive(string entryName, ReaderArchiveOptions options) {
        if (string.IsNullOrWhiteSpace(entryName)) return false;
        string extension = Path.GetExtension(entryName);
        return ArchiveExtensions.Contains(extension)
            || options.ReadInstallerPackages && string.Equals(extension, ".msi", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Whether a member would be read: a nested archive or a supported document.</summary>
    private static bool IsWanted(string entryName, ReaderArchiveOptions options) => IsNestedArchive(entryName, options) || ShouldAttemptRead(entryName);

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

    /// <summary>State shared by every container and member of one read.</summary>
    private sealed class ReadContext {
        internal ReadContext(ArchiveSourceMetadata source, ReaderOptions readerOptions, ArchiveTraversalOptions traversal,
            ReaderArchiveOptions archive, CancellationToken token) {
            Source = source;
            ReaderOptions = readerOptions;
            Traversal = traversal;
            Archive = archive;
            Token = token;
        }

        internal ArchiveSourceMetadata Source { get; }
        internal ReaderOptions ReaderOptions { get; }
        internal ArchiveTraversalOptions Traversal { get; }
        internal ReaderArchiveOptions Archive { get; }
        internal WarningCounter Warnings { get; } = new WarningCounter();
        internal CancellationToken Token { get; }
    }
}
