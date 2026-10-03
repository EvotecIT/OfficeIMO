namespace OfficeIMO.Zip;

/// <summary>
/// Describes a ZIP archive entry.
/// </summary>
public sealed class ZipEntryDescriptor {
    /// <summary>
    /// Entry path inside the ZIP.
    /// </summary>
    public string FullName { get; set; } = string.Empty;

    /// <summary>Original entry name, before display-path normalization.</summary>
    public string RawFullName { get; set; } = string.Empty;

    /// <summary>Zero-based entry index in the original archive, including skipped entries.</summary>
    public int EntryIndex { get; set; }

    /// <summary>
    /// File name component of the entry.
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Indicates whether this entry is a directory.
    /// </summary>
    public bool IsDirectory { get; set; }

    /// <summary>
    /// Path depth calculated from slash segments.
    /// </summary>
    public int Depth { get; set; }

    /// <summary>
    /// Uncompressed length in bytes.
    /// </summary>
    public long UncompressedLength { get; set; }

    /// <summary>
    /// Last write timestamp (UTC) when available.
    /// </summary>
    public DateTime LastWriteUtc { get; set; }
}
