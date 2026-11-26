namespace OfficeIMO.Archives;

/// <summary>An entry in a 7z file table. Listing checks metadata and paths without decoding file payloads.</summary>
public sealed class SevenZipArchiveEntry {
    internal SevenZipArchiveEntry(int ordinal, string name, string relativePath, ArchiveEntryKind kind, long length,
        DateTimeOffset? lastWriteTime, int? folderOrdinal, ArchiveCompression compression) {
        Ordinal = ordinal;
        Name = name;
        RelativePath = relativePath;
        Kind = kind;
        Length = length;
        LastWriteTime = lastWriteTime;
        FolderOrdinal = folderOrdinal;
        Compression = compression;
    }

    /// <summary>Zero-based ordinal in the authored file table.</summary>
    public int Ordinal { get; }
    /// <summary>Exact decoded member name.</summary>
    public string Name { get; }
    /// <summary>Validated portable path with forward slashes.</summary>
    public string RelativePath { get; }
    /// <summary>Member kind. Extraction only publishes regular files.</summary>
    public ArchiveEntryKind Kind { get; }
    /// <summary>Size of the declared uncompressed member stream, or zero when no stream is present.</summary>
    public long Length { get; }
    /// <summary>Authored UTC modification time when representable.</summary>
    public DateTimeOffset? LastWriteTime { get; }
    /// <summary>Zero-based folder ordinal, shared by members of a solid folder. Null for entries without a stream.</summary>
    public int? FolderOrdinal { get; }
    /// <summary>Compression of the containing folder, or None for entries without a stream.</summary>
    public ArchiveCompression Compression { get; }
}
