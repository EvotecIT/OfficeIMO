namespace OfficeIMO.Reader.Toml;

/// <summary>
/// Options for TOML adapter behavior.
/// </summary>
public sealed class TomlReadOptions {
    /// <summary>
    /// Rows per emitted chunk.
    /// </summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>
    /// Include markdown table previews in emitted chunks.
    /// </summary>
    public bool IncludeMarkdown { get; set; } = true;

    /// <summary>
    /// Load settings for the TOML engine: the depth, node, string-length and input-size limits.
    /// <see langword="null"/> uses <see cref="TomlLoadOptions.UntrustedDefaults"/>.
    /// <see cref="ReaderOptions.MaxInputBytes"/> still applies and wins when it is lower.
    /// </summary>
    public TomlLoadOptions? LoadOptions { get; set; }
}
