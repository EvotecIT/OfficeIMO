using OfficeIMO.Archives.SevenZip;

namespace OfficeIMO.Archives;

public static partial class ArchiveReader {
    /// <summary>Lists the validated file table of a 7z archive. Encoded metadata may be decompressed,
    /// but file payloads are not decoded. The source is read from its current position and left open.</summary>
    /// <exception cref="NotSupportedException">The archive uses encryption or an unsupported structure.</exception>
    public static IReadOnlyList<SevenZipArchiveEntry> ListSevenZip(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        return SevenZipArchive.Open(snapshot, limits, cancellationToken).List();
    }

    /// <summary>Reads selected regular files from 7z bytes, verifying the complete containing solid folders
    /// before returning. The caller must leave the input bytes unchanged for the duration of the call.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadSevenZip(byte[] data, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(data);
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        cancellationToken.ThrowIfCancellationRequested();
        ArchiveLimits.Check(data.LongLength, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
        limits.Operation.AdmitInput(data.LongLength);
        using var source = new MemoryStream(data, writable: false);
        return ReadSevenZip(source, limits, cancellationToken);
    }

    /// <summary>Reads selected regular files of a 7z archive from the stream's current position, leaving it open.
    /// Entire selected solid folders are verified, even when only one member is requested.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadSevenZip(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        return ReadSevenZip(snapshot, limits, cancellationToken);
    }

    private static IReadOnlyList<ArchiveEntryData> ReadSevenZip(Stream source, ArchiveLimits limits, CancellationToken token) {
        SevenZipArchive archive = SevenZipArchive.Open(source, limits, token);
        var entries = new List<ArchiveEntryData>();
        archive.Read((path, buffer, offset, length) => {
            var owned = new byte[length];
            Buffer.BlockCopy(buffer, offset, owned, 0, length);
            entries.Add(new ArchiveEntryData(path, owned));
        }, token);
        return Array.AsReadOnly(entries.OrderBy(entry => entry.RelativePath, StringComparer.Ordinal).ToArray());
    }
}
