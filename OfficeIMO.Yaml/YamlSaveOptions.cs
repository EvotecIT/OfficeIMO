namespace OfficeIMO.Yaml;

/// <summary>
/// Settings for saving a <see cref="YamlDocument"/>. The default save changes only edited text; everything else is
/// copied from the source byte for byte.
/// </summary>
public sealed class YamlSaveOptions {
    /// <summary>
    /// Gets or sets whether to write the whole document anew through the emitter instead of copying untouched text, to
    /// normalize indentation. Values, scalar spellings and styles, anchors, tags, comments and the blank lines above
    /// entries are kept; multi-line plain and folded scalars are laid out again, and comments inside flow collections
    /// are dropped. Documents loaded with <see cref="YamlDialect.Unity"/> cannot be reformatted.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    public bool Reformat { get; set; }

    /// <summary>
    /// Gets or sets the layout a reformat uses, or <see langword="null"/> (the default) for the layout inferred from the
    /// document itself.
    /// </summary>
    public YamlWriterOptions? WriterOptions { get; set; }

    /// <summary>
    /// The quotes a reformat gives quoted scalars (<see cref="YamlStyle.SingleQuoted"/> or
    /// <see cref="YamlStyle.DoubleQuoted"/>), or <see cref="YamlStyle.None"/> to keep each scalar's own.
    /// <see cref="YamlFormatter"/> sets it.
    /// </summary>
    internal YamlStyle QuoteStyle { get; set; }
}
