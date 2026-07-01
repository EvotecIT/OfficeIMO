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
/// Markdig-style abbreviation definition collected while parsing, such as
/// <c>*[HTML]: Hyper Text Markup Language</c>.
/// </summary>
public sealed class MarkdownAbbreviationDefinition {
    /// <summary>Source spelling that should be expanded when it appears as a standalone inline token.</summary>
    public string Label { get; }

    /// <summary>Title rendered on the resulting <c>&lt;abbr&gt;</c> element.</summary>
    public string Title { get; }

    /// <summary>Source span for the entire abbreviation definition, when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Source span for the abbreviation label token, when available.</summary>
    public MarkdownSourceSpan? LabelSourceSpan { get; }

    /// <summary>Source span for the opening <c>*[</c> marker, when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; }

    /// <summary>Source span for the closing <c>]:</c> marker, when available.</summary>
    public MarkdownSourceSpan? SeparatorMarkerSourceSpan { get; }

    /// <summary>Source span for the title token, when available.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>Whether the definition was authored as leading content inside a list item.</summary>
    public bool IsListItemDefinition { get; }

    /// <summary>Creates an abbreviation definition descriptor.</summary>
    public MarkdownAbbreviationDefinition(
        string label,
        string title,
        MarkdownSourceSpan? sourceSpan = null,
        MarkdownSourceSpan? labelSourceSpan = null,
        MarkdownSourceSpan? titleSourceSpan = null,
        MarkdownSourceSpan? openingMarkerSourceSpan = null,
        MarkdownSourceSpan? separatorMarkerSourceSpan = null,
        bool isListItemDefinition = false) {
        Label = label ?? string.Empty;
        Title = title ?? string.Empty;
        SourceSpan = sourceSpan;
        LabelSourceSpan = labelSourceSpan;
        OpeningMarkerSourceSpan = openingMarkerSourceSpan;
        SeparatorMarkerSourceSpan = separatorMarkerSourceSpan;
        TitleSourceSpan = titleSourceSpan;
        IsListItemDefinition = isListItemDefinition;
    }
}

/// <summary>
/// Mutable per-parse state shared across block and inline parsers.
/// </summary>
public sealed class MarkdownReaderState {
    /// <summary>Reference-style link definitions collected while parsing.</summary>
    internal Dictionary<string, MarkdownReferenceLinkDefinition> LinkRefs { get; } = new Dictionary<string, MarkdownReferenceLinkDefinition>(System.StringComparer.OrdinalIgnoreCase);
    /// <summary>Case-sensitive abbreviation definitions collected while parsing. Later definitions replace earlier ones, matching Markdig.</summary>
    internal Dictionary<string, MarkdownAbbreviationDefinition> Abbreviations { get; } = new Dictionary<string, MarkdownAbbreviationDefinition>(System.StringComparer.Ordinal);
    internal int SourceLineOffset { get; set; }
    internal MarkdownSourceTextMap? SourceTextMap { get; set; }
    internal int ListMarkerIndentOffset { get; set; }
    internal bool SuppressBlockGenericAttributes { get; set; }
    internal bool SuppressHeadingGenericAttributes { get; set; }
    internal bool IsMarkdigDefinitionListBody { get; set; }
    internal MarkdownPendingGenericAttributeBlock? PendingGenericAttributeBlock { get; set; }
    internal HashSet<int> LazyQuoteContinuationLines { get; } = new HashSet<int>();
    internal HashSet<int> QuoteContainerLines { get; } = new HashSet<int>();
    internal HashSet<int> SuppressedSetextHeadingUnderlineLines { get; } = new HashSet<int>();
    internal HashSet<int> SuppressedParagraphGenericAttributeStartLines { get; } = new HashSet<int>();
    internal IReadOnlyList<int>? SourceLineAbsoluteNumbers { get; set; }
}

internal sealed class MarkdownPendingGenericAttributeBlock {
    internal MarkdownPendingGenericAttributeBlock(
        MarkdownAttributeSet attributes,
        string sourceText,
        MarkdownSourceSpan sourceSpan) {
        Attributes = attributes ?? MarkdownAttributeSet.Empty;
        SourceText = sourceText ?? string.Empty;
        SourceSpan = sourceSpan;
    }

    internal MarkdownAttributeSet Attributes { get; }
    internal string SourceText { get; }
    internal MarkdownSourceSpan SourceSpan { get; }
}
