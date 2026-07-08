namespace OfficeIMO.Markdown;

/// <summary>
/// Controls how parse-owned abbreviation definitions are emitted by <see cref="MarkdownDoc.ToMarkdown(MarkdownWriteOptions?)"/>.
/// </summary>
public enum MarkdownAbbreviationDefinitionRenderingMode {
    /// <summary>Preserve abbreviation definition syntax in the generated Markdown.</summary>
    Preserve,
    /// <summary>Do not emit parse-owned abbreviation definitions.</summary>
    Omit
}
