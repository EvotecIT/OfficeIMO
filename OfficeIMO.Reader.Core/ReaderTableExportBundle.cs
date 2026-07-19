namespace OfficeIMO.Reader;

/// <summary>
/// Deterministic export payloads for a discovered reader table.
/// </summary>
public sealed class ReaderTableExportBundle {
    /// <summary>
    /// Stable table export identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Filesystem-safe stem suitable for CSV, Markdown, and JSON sidecar files.
    /// </summary>
    public string FileNamePrefix { get; set; } = string.Empty;

    /// <summary>
    /// Source table model.
    /// </summary>
    public ReaderTable Table { get; set; } = new ReaderTable();

    /// <summary>
    /// Deterministic CSV representation.
    /// </summary>
    public string Csv { get; set; } = string.Empty;

    /// <summary>
    /// Deterministic GitHub-style Markdown table representation.
    /// </summary>
    public string Markdown { get; set; } = string.Empty;

    /// <summary>
    /// Deterministic JSON representation.
    /// </summary>
    public string Json { get; set; } = string.Empty;
}
