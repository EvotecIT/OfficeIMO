namespace OfficeIMO.Reader.Json;

/// <summary>
/// Options for JSON adapter behavior.
/// </summary>
public sealed class JsonReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Maximum traversal depth.
    /// </summary>
    public int MaxDepth { get; set; } = 32;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
