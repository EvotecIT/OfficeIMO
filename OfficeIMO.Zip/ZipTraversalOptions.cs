namespace OfficeIMO.Zip;

/// <summary>
/// Controls safe ZIP traversal behavior.
/// </summary>
public sealed class ZipTraversalOptions {
    /// <summary>
    /// Maximum number of entries to emit.
    /// </summary>
    public int MaxEntries { get; set; } = 5000;

    /// <summary>
    /// Maximum path depth for entries.
    /// </summary>
    public int MaxDepth { get; set; } = 16;

    /// <summary>
    /// Maximum total uncompressed bytes across emitted entries.
    /// </summary>
    public long? MaxTotalUncompressedBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>
    /// Optional maximum uncompressed size for a single file entry.
    /// </summary>
    public long? MaxEntryUncompressedBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>
    /// Optional maximum uncompressed/compressed ratio for a single file entry.
    /// Helps reject high-expansion entries from suspicious archives.
    /// </summary>
    public double? MaxCompressionRatio { get; set; } = 300;

    /// <summary>
    /// When true, includes directory entries.
    /// </summary>
    public bool IncludeDirectoryEntries { get; set; }

    /// <summary>
    /// When true, entry order is deterministic (ordinal by FullName).
    /// </summary>
    public bool DeterministicOrder { get; set; } = true;
}
