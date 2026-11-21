namespace OfficeIMO.Reader.Toml;

/// <summary>
/// Options for TOML adapter behavior.
/// </summary>
/// <remarks>
/// Each limit starts from <see cref="LoadOptions"/>, or from <see cref="TomlLoadOptions.UntrustedDefaults"/> when that is
/// <see langword="null"/>. A limit set on this object replaces the value from that base. Exceeding a limit throws
/// <see cref="ReaderResourceLimitException"/> named after the property on this object.
/// </remarks>
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
    /// Maximum nesting of tables, arrays and inline tables, counting each key of a dotted key or header.
    /// <see langword="null"/> keeps the base value (64 with <see cref="TomlLoadOptions.UntrustedDefaults"/>).
    /// </summary>
    public int? MaxDepth { get; set; }

    /// <summary>
    /// Maximum number of tables, arrays and values in the document.
    /// <see langword="null"/> keeps the base value (200,000 with <see cref="TomlLoadOptions.UntrustedDefaults"/>).
    /// </summary>
    public int? MaxNodes { get; set; }

    /// <summary>
    /// Maximum source length of one string or key in UTF-16 code units, including its quotes. Maps to
    /// <see cref="TomlLoadOptions.MaxStringLength"/>. <see langword="null"/> keeps the base value
    /// (1,048,576 with <see cref="TomlLoadOptions.UntrustedDefaults"/>).
    /// </summary>
    public int? MaxScalarLength { get; set; }

    /// <summary>
    /// Load settings for the TOML engine: the depth, node, string-length and input-size limits.
    /// <see langword="null"/> uses <see cref="TomlLoadOptions.UntrustedDefaults"/>. <see cref="MaxDepth"/>,
    /// <see cref="MaxNodes"/> and <see cref="MaxScalarLength"/> override the matching values when set.
    /// <see cref="ReaderOptions.MaxInputBytes"/> still applies and wins when it is lower.
    /// </summary>
    public TomlLoadOptions? LoadOptions { get; set; }
}
