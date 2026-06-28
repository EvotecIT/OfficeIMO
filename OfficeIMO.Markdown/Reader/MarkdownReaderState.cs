namespace OfficeIMO.Markdown;

/// <summary>
/// Reference-style link definition collected while parsing, such as <c>[label]: https://example.com "Title"</c>.
/// </summary>
public sealed class MarkdownReferenceLinkDefinition {
    /// <summary>Normalized reference label used for matching reference-style links and images.</summary>
    public string Label { get; }

    /// <summary>Resolved destination URL after reader profile URL handling.</summary>
    public string Url { get; }

    /// <summary>Optional definition title.</summary>
    public string? Title { get; }

    /// <summary>Source span for the entire reference-style definition, when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Source span for the definition label token, when available.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>[</c> marker before the definition label, when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the <c>]:</c> marker after the definition label, when available.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the destination token, when available.</summary>
    public MarkdownSourceSpan? UrlSourceSpan { get; }

    /// <summary>Source span for the optional title token, when available.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>
    /// Creates a reference-style link definition descriptor.
    /// </summary>
    public MarkdownReferenceLinkDefinition(
        string label,
        string url,
        string? title,
        MarkdownSourceSpan? sourceSpan = null,
        MarkdownSourceSpan? labelSourceSpan = null,
        MarkdownSourceSpan? urlSourceSpan = null,
        MarkdownSourceSpan? titleSourceSpan = null,
        MarkdownSourceSpan? openingMarkerSourceSpan = null,
        MarkdownSourceSpan? separatorMarkerSourceSpan = null) {
        Label = label ?? string.Empty;
        Url = url ?? string.Empty;
        Title = title;
        SourceSpan = sourceSpan;
        LabelSourceSpan = labelSourceSpan;
        OpeningMarkerSourceSpan = openingMarkerSourceSpan;
        SeparatorMarkerSourceSpan = separatorMarkerSourceSpan;
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
    internal int ListMarkerIndentOffset { get; set; }
    internal HashSet<int> LazyQuoteContinuationLines { get; } = new HashSet<int>();
}
