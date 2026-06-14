namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// Options for YAML adapter behavior.
/// </summary>
public sealed class YamlReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Maximum node traversal depth.
    /// </summary>
    public int MaxDepth { get; set; } = 32;

    /// <summary>
    /// Maximum number of YAML nodes visited before traversal emits a node-limit row and stops.
    /// </summary>
    public int MaxNodes { get; set; } = 20_000;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
