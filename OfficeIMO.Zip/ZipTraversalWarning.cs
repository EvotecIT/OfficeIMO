namespace OfficeIMO.Zip;

/// <summary>
/// Describes a warning produced while traversing ZIP entries.
/// </summary>
public sealed class ZipTraversalWarning {
    /// <summary>
    /// Entry path related to the warning, when available.
    /// </summary>
    public string EntryPath { get; set; } = string.Empty;

    /// <summary>
    /// Warning message.
    /// </summary>
    public string Warning { get; set; } = string.Empty;
}
