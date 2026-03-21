namespace OfficeIMO.Markdown;

internal sealed class MarkdownReferenceLinkDefinition {
    public string Label { get; }
    public string Url { get; }
    public string? Title { get; }
    public MarkdownSourceSpan? LabelSourceSpan { get; }
    public MarkdownSourceSpan? UrlSourceSpan { get; }
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    public MarkdownReferenceLinkDefinition(
        string label,
        string url,
        string? title,
        MarkdownSourceSpan? labelSourceSpan = null,
        MarkdownSourceSpan? urlSourceSpan = null,
        MarkdownSourceSpan? titleSourceSpan = null) {
        Label = label ?? string.Empty;
        Url = url ?? string.Empty;
        Title = title;
        LabelSourceSpan = labelSourceSpan;
        UrlSourceSpan = urlSourceSpan;
        TitleSourceSpan = titleSourceSpan;
    }
}

/// <summary>
/// Mutable per-parse state shared across block and inline parsers.
/// </summary>
public sealed class MarkdownReaderState {
    /// <summary>Reference-style link definitions collected while parsing.</summary>
    internal Dictionary<string, MarkdownReferenceLinkDefinition> LinkRefs { get; } = new Dictionary<string, MarkdownReferenceLinkDefinition>(System.StringComparer.OrdinalIgnoreCase);
    internal int SourceLineOffset { get; set; }
    internal MarkdownSourceTextMap? SourceTextMap { get; set; }
}
