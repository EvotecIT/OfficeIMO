namespace OfficeIMO.Reader;

/// <summary>
/// Deterministic export payloads for a discovered reader visual.
/// </summary>
public sealed class ReaderVisualExportBundle {
    /// <summary>
    /// Stable visual export identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Filesystem-safe stem suitable for visual payload and JSON sidecar files.
    /// </summary>
    public string FileNamePrefix { get; set; } = string.Empty;

    /// <summary>
    /// File extension for the source visual payload.
    /// </summary>
    public string PayloadExtension { get; set; } = ".txt";

    /// <summary>
    /// Source visual model.
    /// </summary>
    public ReaderVisual Visual { get; set; } = new ReaderVisual();

    /// <summary>
    /// Raw source visual payload.
    /// </summary>
    public string Payload { get; set; } = string.Empty;

    /// <summary>
    /// Deterministic JSON representation.
    /// </summary>
    public string Json { get; set; } = string.Empty;
}
