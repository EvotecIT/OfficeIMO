using OfficeIMO.Archives.Rar5;

namespace OfficeIMO.Archives;

public static partial class ArchiveReader {
    /// <summary>Lists the authoritative file headers of a RAR5 archive in file order. The source remains open.</summary>
    public static IReadOnlyList<Rar5ArchiveEntry> ListRar5(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) => InspectRar5(source, options, cancellationToken).Entries;

    /// <summary>Reads the authoritative file table and verifies the archive comment, if any. The source remains open.</summary>
    public static Rar5ArchiveInfo InspectRar5(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        cancellationToken.ThrowIfCancellationRequested();
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        return Rar5Archive.Open(snapshot, limits, cancellationToken).Describe(1);
    }

    /// <summary>Reads the file table and optional comment of a complete RAR5 set supplied in volume order.
    /// Each stream is read from its current position to EOF and left open.</summary>
    public static Rar5ArchiveInfo InspectRar5Volumes(IReadOnlyList<Stream> volumes, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        WithRar5Volumes(volumes, options, cancellationToken, archive => archive.Describe(volumes.Count));

    /// <summary>Returns selected stored RAR5 files in primary header order after verifying their checksums.
    /// The caller must keep source bytes unchanged for the duration of the read.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadRar5(byte[] data, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(data);
        cancellationToken.ThrowIfCancellationRequested();
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        ArchiveLimits.Check(data.LongLength, limits.Input, nameof(ArchiveReadOptions.MaxInputBytes));
        limits.Operation.AdmitInput(data.LongLength);
        using var source = new MemoryStream(data, writable: false);
        return ReadRar5(Rar5Archive.Open(source, limits, cancellationToken), cancellationToken);
    }

    /// <summary>Reads and verifies selected stored RAR5 entries from the stream's current position to EOF, leaving it open.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadRar5(Stream source, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(source);
        cancellationToken.ThrowIfCancellationRequested();
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        using Stream snapshot = Snapshot(source, limits, cancellationToken);
        return ReadRar5(Rar5Archive.Open(snapshot, limits, cancellationToken), cancellationToken);
    }

    /// <summary>Returns the verified stored files of a complete RAR5 set supplied in volume order.
    /// Each stream is read from its current position to EOF and left open.</summary>
    public static IReadOnlyList<ArchiveEntryData> ReadRar5Volumes(IReadOnlyList<Stream> volumes, ArchiveReadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        WithRar5Volumes(volumes, options, cancellationToken, archive => ReadRar5(archive, cancellationToken));

    private static IReadOnlyList<ArchiveEntryData> ReadRar5(Rar5Archive archive, CancellationToken token) {
        var members = new List<ArchiveEntryData>();
        archive.Read((path, data) => members.Add(new ArchiveEntryData(path, data)), token);
        return members.AsReadOnly();
    }

    private static T WithRar5Volumes<T>(IReadOnlyList<Stream> volumes, ArchiveReadOptions? options, CancellationToken token,
        Func<Rar5Archive, T> consume) {
        ArgumentNullException.ThrowIfNull(volumes);
        token.ThrowIfCancellationRequested();
        int count = volumes.Count;
        if (count == 0) throw new ArgumentException("Supply at least one RAR5 volume.", nameof(volumes));
        ArchiveLimits.Check(count, Rar5Archive.MaxVolumeCount, "MaxVolumeCount");
        ArchiveLimits limits = (options ?? new ArchiveReadOptions()).Limits();
        var snapshots = new List<Stream>(count);
        long total = 0;
        try {
            for (int i = 0; i < count; i++) {
                Stream input = volumes[i] ?? throw new ArgumentException("A RAR5 volume stream cannot be null.", nameof(volumes));
                Stream snapshot = Snapshot(input, limits, token, limits.Input - total);
                snapshots.Add(snapshot);
                total += snapshot.Length;
            }
            return consume(Rar5Archive.Open(snapshots, limits, token));
        } finally {
            foreach (Stream snapshot in snapshots) snapshot.Dispose();
        }
    }
}
