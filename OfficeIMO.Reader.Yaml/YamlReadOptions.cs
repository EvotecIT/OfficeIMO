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
    /// Maximum node depth accepted by the streaming preflight before the representation model is loaded.
    /// </summary>
    public int MaxDepth { get; set; } = 32;

    /// <summary>
    /// Maximum number of YAML nodes accepted by the streaming preflight before the representation model is loaded.
    /// </summary>
    public int MaxNodes { get; set; } = 20_000;

    /// <summary>
    /// Maximum number of YAML parse events accepted before the representation model is loaded.
    /// </summary>
    public int MaxParseEvents { get; set; } = 100_000;

    /// <summary>
    /// Maximum scalar value length accepted before the representation model is loaded.
    /// </summary>
    public int MaxScalarLength { get; set; } = 1_048_576;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;
}
