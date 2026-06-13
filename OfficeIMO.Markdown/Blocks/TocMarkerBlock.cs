namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown table-of-contents marker that preserves a semantic TOC placeholder when rendering back to Markdown.
/// </summary>
public sealed class TocMarkerBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Minimum heading level included by the table of contents.</summary>
    public int MinLevel { get; set; } = TocOptions.DefaultMinLevel;

    /// <summary>Maximum heading level included by the table of contents.</summary>
    public int MaxLevel { get; set; } = TocOptions.DefaultMaxLevel;

    /// <summary>When true, output formats may render a title above the generated table of contents.</summary>
    public bool IncludeTitle { get; set; } = true;

    /// <summary>Title text for output formats that render TOC chrome.</summary>
    public string Title { get; set; } = "Table of Contents";

    /// <summary>Heading level for the title when an output format renders it as a heading.</summary>
    public int TitleLevel { get; set; } = TocOptions.DefaultTitleLevel;

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        var sb = new StringBuilder("[TOC");
        sb.Append(" min=").Append(ClampTocLevel(MinLevel));
        sb.Append(" max=").Append(ClampTocLevel(MaxLevel));
        if (IncludeTitle && !string.IsNullOrWhiteSpace(Title)) {
            sb.Append(" title=\"").Append(EscapeAttributeValue(Title.Trim())).Append('"');
            sb.Append(" titleLevel=").Append(ClampTitleLevel(TitleLevel));
        }

        sb.Append(']');
        return sb.ToString();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() => string.Empty;

    /// <inheritdoc />
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.TocPlaceholder, span, ((IMarkdownBlock)this).RenderMarkdown(), associatedObject: this);

    private static int ClampTocLevel(int level) => level < 1 ? 1 : (level > 9 ? 9 : level);

    private static int ClampTitleLevel(int level) => level < 1 ? 1 : (level > 6 ? 6 : level);

    private static string EscapeAttributeValue(string value) =>
        value.Replace("\\", "\\\\").Replace("\"", "\\\"");
}
