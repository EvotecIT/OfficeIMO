namespace OfficeIMO.Archives;

/// <summary>Settings shared by the ZIP and TAR writers. The defaults produce identical bytes for identical input
/// on every platform and target framework.</summary>
public abstract class ArchiveWriterOptions {
    /// <summary>The time written for every entry. When null, an entry's own time is used, and entries without one get
    /// the format's fixed default (1980-01-01 for ZIP, the Unix epoch for TAR).</summary>
    public DateTimeOffset? Timestamp { get; set; }
    /// <summary>When true and <see cref="Timestamp"/> is set, only entry times later than it are replaced, matching the
    /// reproducible-builds clamping convention. Entries without a time still get <see cref="Timestamp"/>.</summary>
    public bool ClampToTimestamp { get; set; }
    /// <summary>When true, the directory helpers record each file's last write time. Defaults to false so that
    /// checkout and copy times do not change the archive bytes.</summary>
    public bool PreserveFileTimes { get; set; }
    /// <summary>Maximum path segments per member. Defaults to 64, the extraction default.</summary>
    public int MaxPathDepth { get; set; } = 64;
    /// <summary>Maximum characters per member path. Defaults to 1024, the extraction default.</summary>
    public int MaxPathLength { get; set; } = 1024;

    internal DateTimeOffset? Resolve(DateTimeOffset? entryTime) {
        if (!Timestamp.HasValue) return entryTime;
        if (ClampToTimestamp && entryTime.HasValue && entryTime.Value <= Timestamp.Value) return entryTime;
        return Timestamp;
    }

    internal ArchivePaths CreatePaths() => new(new ArchiveLimits(new ArchiveExtractionOptions {
        MaxPathDepth = MaxPathDepth, MaxPathLength = MaxPathLength
    }));
}
