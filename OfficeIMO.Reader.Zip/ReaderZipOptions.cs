namespace OfficeIMO.Reader.Zip;

/// <summary>
/// Controls nested archive behavior for ZIP ingestion.
/// </summary>
public sealed class ReaderZipOptions {
    /// <summary>
    /// When true, nested .zip entries are traversed recursively.
    /// </summary>
    public bool ReadNestedZipEntries { get; set; } = true;

    /// <summary>
    /// Maximum recursion depth for nested archives. 0 means top-level only.
    /// </summary>
    public int MaxNestedDepth { get; set; } = 2;

    /// <summary>
    /// Optional max byte size for nested ZIP entry materialization.
    /// </summary>
    public long? MaxNestedArchiveBytes { get; set; } = 64L * 1024 * 1024;
}
