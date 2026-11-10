namespace OfficeIMO.Reader.Yaml;

/// <summary>
/// Options for YAML adapter behavior.
/// </summary>
public sealed class YamlReadOptions {
    /// <summary>
    /// Rows per emitted chunk. Each document of a multi-document stream starts a new chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;

    /// <summary>
    /// Load settings for the YAML engine: schema, duplicate-key policy and the depth, node, alias-expansion, scalar
    /// length and input-size limits. <see langword="null"/> uses the <see cref="YamlLoadOptions"/> defaults.
    /// <see cref="ReaderOptions.MaxInputBytes"/> still applies and wins when it is lower.
    /// </summary>
    public YamlLoadOptions? LoadOptions { get; set; }
}
