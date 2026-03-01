namespace OfficeIMO.Zip;

/// <summary>
/// ZIP traversal output including accepted entries and warnings.
/// </summary>
public sealed class ZipTraversalResult {
    /// <summary>
    /// Accepted entries that passed traversal rules.
    /// </summary>
    public IReadOnlyList<ZipEntryDescriptor> Entries { get; set; } = Array.Empty<ZipEntryDescriptor>();

    /// <summary>
    /// Warnings produced while evaluating entries.
    /// </summary>
    public IReadOnlyList<ZipTraversalWarning> Warnings { get; set; } = Array.Empty<ZipTraversalWarning>();

    /// <summary>
    /// Total uncompressed bytes across accepted file entries.
    /// </summary>
    public long TotalUncompressedBytes { get; set; }

    /// <summary>
    /// Number of archive entries visited.
    /// </summary>
    public int EntriesVisited { get; set; }
}
