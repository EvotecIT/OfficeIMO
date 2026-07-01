namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown table-of-contents marker that preserves a semantic TOC placeholder when rendering back to Markdown.
/// </summary>
public sealed class TocMarkerBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IContextualHtmlMarkdownBlock {
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
    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        var options = new TocOptions {
            IncludeTitle = IncludeTitle,
            Title = Title,
            TitleLevel = TitleLevel,
            MinLevel = ClampTocLevel(MinLevel),
            MaxLevel = ClampTocLevel(MaxLevel),
            RequireTopLevel = false
        };

        if (options.MaxLevel < options.MinLevel) {
            options.MaxLevel = options.MinLevel;
        }

        int blockIndex = context.GetBlockIndex(this);
        var entries = context.BuildTocEntries(blockIndex, options);
        if (entries.Count == 0) {
            return string.Empty;
        }

        var overridden = context.Options.TocHtmlRenderer?.Invoke(options, entries, context.Options);
        if (overridden != null) {
            return overridden;
        }

        var toc = new TocBlock {
            Ordered = options.Ordered,
            NormalizeLevels = options.NormalizeToMinLevel,
            IncludeTitle = false,
            MinLevel = options.MinLevel,
            MaxLevel = options.MaxLevel,
            RequireTopLevel = false
        };

        for (int i = 0; i < entries.Count; i++) {
            toc.Entries.Add(entries[i]);
        }

        string listHtml = ((IMarkdownBlock)toc).RenderHtml();
        if (!IncludeTitle || string.IsNullOrWhiteSpace(Title)) {
            return listHtml;
        }

        int titleLevel = ClampTitleLevel(TitleLevel);
        return new StringBuilder()
            .Append("<h").Append(titleLevel).Append('>')
            .Append(HtmlTextEncoder.Encode(Title.Trim(), context.Options))
            .Append("</h").Append(titleLevel).Append('>')
            .Append(listHtml)
            .ToString();
    }

    /// <inheritdoc />
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.TocPlaceholder, span, ((IMarkdownBlock)this).RenderMarkdown(), associatedObject: this);

    private static int ClampTocLevel(int level) => level < 1 ? 1 : (level > 9 ? 9 : level);

    private static int ClampTitleLevel(int level) => level < 1 ? 1 : (level > 6 ? 6 : level);

    private static string EscapeAttributeValue(string value) =>
        value.Replace("\\", "\\\\").Replace("\"", "\\\"");
}
