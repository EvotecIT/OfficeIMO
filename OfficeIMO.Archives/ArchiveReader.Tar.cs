using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

public static partial class ArchiveReader {
    /// <summary>Reads the selected regular files of a TAR or gzip-compressed TAR (.tar.gz) from bytes, with the same path
    /// safety, collision and limit checks as extraction. Links, devices and directories are not returned; use
    /// <see cref="ListTar(Stream, ArchiveReadOptions?, CancellationToken)"/> for their metadata.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The archive is corrupt, truncated or unsafe.</exception>
    /// <exception cref="NotSupportedException">The input is not TAR, or uses sparse or multi-volume members.</exception>
    public static IReadOnlyList<ArchiveEntryData> ReadTar(byte[] bytes, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(bytes);
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        ArchiveLimits.Check(bytes.LongLength, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
        limits.Operation.AdmitInput(bytes.LongLength);
        using var source = new MemoryStream(bytes, writable: false);
        return ReadTar(source, limits, cancellationToken);
    }

    /// <summary>Reads the selected regular files of a TAR or .tar.gz from the stream's current position to EOF.
    /// Supports nonseekable streams, leaves the caller's stream open and advances it by the bytes read.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The archive is corrupt, truncated or unsafe.</exception>
    /// <exception cref="NotSupportedException">The input is not TAR, or uses sparse or multi-volume members.</exception>
    public static IReadOnlyList<ArchiveEntryData> ReadTar(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        return ReadTar(snapshot, limits, cancellationToken);
    }

    /// <summary>Lists every member of a TAR or .tar.gz, including directories, symbolic and hard links with their authored
    /// targets, and device records, without reading file content or validating names as output paths.
    /// Reads from the stream's current position to EOF and leaves the stream open.</summary>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">The archive is corrupt or truncated.</exception>
    /// <exception cref="NotSupportedException">The input is not TAR, or uses sparse or multi-volume members.</exception>
    public static IReadOnlyList<TarArchiveEntry> ListTar(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        var limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        using TarSource tar = TarSource.Open(snapshot, limits, cancellationToken);
        return Array.AsReadOnly(TarArchiveContent.ReadDirectory(tar.Stream, limits, cancellationToken).ToArray());
    }

    private static IReadOnlyList<ArchiveEntryData> ReadTar(Stream source, ArchiveLimits limits, CancellationToken token) {
        using TarSource tar = TarSource.Open(source, limits, token);
        var files = new List<ArchiveEntryData>();
        foreach (var (entry, path) in tar.SelectFiles(limits, token, rejectLinks: false)) {
            ArchiveLimits.Check(entry.Length, Math.Min(limits.Entry, int.MaxValue), nameof(ArchiveReadOptions.MaxEntryBytes));
            var data = new byte[entry.Length];
            TarArchiveContent.ReadExactly(tar.Stream, entry.DataOffset, data);
            files.Add(new ArchiveEntryData(path, data));
        }
        token.ThrowIfCancellationRequested();
        return Array.AsReadOnly(files.OrderBy(file => file.RelativePath, StringComparer.Ordinal).ToArray());
    }

    /// <summary>Copies a readable stream from its current position to EOF into a seekable snapshot under the input limit.</summary>
    internal static Stream Snapshot(Stream source, ArchiveLimits limits, CancellationToken token) {
        if (!source.CanRead) throw new ArgumentException("A readable stream is required.", nameof(source));
        var snapshot = new ArchiveSpillStream();
        try {
            var buffer = new byte[81920];
            limits.Operation.AdmitInput(0);
            while (true) {
                token.ThrowIfCancellationRequested();
                long remaining = limits.Input - snapshot.Length;
                int read = source.Read(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1);
                if (read == 0) break;
                if (read > remaining) throw ArchiveLimits.Exceeded(nameof(ArchiveReadOptions.MaxInputBytes), limits.Input);
                limits.Operation.AdmitInput(read);
                snapshot.Write(buffer, 0, read);
            }
            snapshot.Position = 0;
            return snapshot;
        } catch { snapshot.Dispose(); throw; }
    }
}

/// <summary>A seekable, uncompressed TAR stream over a plain or gzip-compressed input.</summary>
internal sealed class TarSource : IDisposable {
    private readonly ArchiveSpillStream? _decompressed;
    private TarSource(Stream stream, ArchiveSpillStream? decompressed) { Stream = stream; _decompressed = decompressed; }
    internal Stream Stream { get; }
    internal bool Compressed => _decompressed != null;

    /// <summary>Opens a seekable input positioned at its start. Gzip input is decompressed into private staging, bounded by
    /// MaxExpandedBytes plus 8 KiB of header, extended-header and padding allowance per permitted entry and 1 MiB for the
    /// end-of-archive record.</summary>
    internal static TarSource Open(Stream source, ArchiveLimits limits, CancellationToken token) {
        source.Position = 0;
        var head = new byte[1024];
        int length = Fill(source, head);
        source.Position = 0;
        if (GZipContent.HasSignature(head, length)) {
            long cap = Saturating(Saturating(limits.Expanded, 8192L * limits.Entries), 1024L * 1024);
            var staged = new ArchiveSpillStream();
            try {
                GZipContent.Decompress(source, staged, cap, limits, nameof(ArchiveExtractionOptions.MaxExpandedBytes), token);
                staged.Position = 0;
                length = Fill(staged, head);
                staged.Position = 0;
                if (!ArchiveExtractor.LooksLikeTar(head, length)) throw new NotSupportedException("The gzip content is not a TAR archive.");
                return new TarSource(staged, staged);
            } catch { staged.Dispose(); throw; }
        }
        if (!ArchiveExtractor.LooksLikeTar(head, length)) throw new NotSupportedException("A TAR or gzip-compressed TAR archive is required.");
        return new TarSource(source, null);
    }

    /// <summary>Validates every member path, then returns the selected regular files. Selected links and special files
    /// either fail (extraction) or are skipped (in-memory reads).</summary>
    internal List<(TarArchiveEntry Entry, string Path)> SelectFiles(ArchiveLimits limits, CancellationToken token, bool rejectLinks) {
        List<TarArchiveEntry> entries = TarArchiveContent.ReadDirectory(Stream, limits, token);
        var paths = new ArchivePaths(limits);
        var selected = new List<(TarArchiveEntry, string)>();
        long declared = 0;
        foreach (TarArchiveEntry entry in entries) {
            token.ThrowIfCancellationRequested();
            string? name = TarArchiveContent.RelativeName(entry.Name);
            if (name == null) {
                if (entry.Kind == ArchiveEntryKind.Directory) continue;
                throw new InvalidDataException("A TAR member has an empty name.");
            }
            if (entry.Kind == ArchiveEntryKind.Directory) { paths.AddDirectory(name); continue; }
            string path = paths.Add(name);
            if (!limits.Select(path)) continue;
            if (entry.Kind != ArchiveEntryKind.File) {
                if (rejectLinks) throw new InvalidDataException("TAR member '" + path + "' is a " + Describe(entry.Kind)
                    + "; extraction does not create links or special files. Exclude it with IncludeFile.");
                continue;
            }
            ArchiveLimits.Add(ref declared, entry.Length, limits.Expanded, nameof(ArchiveExtractionOptions.MaxExpandedBytes));
            selected.Add((entry, path));
        }
        foreach (var (entry, _) in selected) limits.Operation.ReserveRead(entry.Length, Compressed ? 0 : entry.Length, entry.Length);
        return selected;
    }

    private static string Describe(ArchiveEntryKind kind) => kind == ArchiveEntryKind.SymbolicLink ? "symbolic link"
        : kind == ArchiveEntryKind.HardLink ? "hard link" : "device or FIFO";

    private static long Saturating(long left, long right) => left > long.MaxValue - right ? long.MaxValue : left + right;

    private static int Fill(Stream source, byte[] buffer) {
        int total = 0, read;
        while (total < buffer.Length && (read = source.Read(buffer, total, buffer.Length - total)) > 0) total += read;
        return total;
    }

    public void Dispose() => _decompressed?.Dispose();
}
