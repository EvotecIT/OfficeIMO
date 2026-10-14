using System;
using System.Collections.Generic;
namespace OfficeIMO.Archives;

/// <summary>
/// ZIP traversal output including accepted entries and warnings.
/// </summary>
public sealed class ArchiveTraversalResult {
    /// <summary>
    /// Accepted entries that passed traversal rules.
    /// </summary>
    public IReadOnlyList<ArchiveEntryDescriptor> Entries { get; set; } = Array.Empty<ArchiveEntryDescriptor>();

    /// <summary>
    /// Warnings produced while evaluating entries.
    /// </summary>
    public IReadOnlyList<ArchiveTraversalWarning> Warnings { get; set; } = Array.Empty<ArchiveTraversalWarning>();

    /// <summary>
    /// Total uncompressed bytes across accepted file entries.
    /// </summary>
    public long TotalUncompressedBytes { get; set; }

    /// <summary>
    /// Number of archive entries visited.
    /// </summary>
    public int EntriesVisited { get; set; }
}
