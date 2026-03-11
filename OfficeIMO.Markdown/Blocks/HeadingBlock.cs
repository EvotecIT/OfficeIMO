namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown heading (ATX) block, levels 1–6.
/// </summary>
public sealed class HeadingBlock : IMarkdownBlock, ISyntaxMarkdownBlock, IContextualHtmlMarkdownBlock {
    /// <summary>Heading level constrained to [1,6].</summary>
    public int Level { get; }
    /// <summary>Inline content owned by this heading.</summary>
    public InlineSequence Inlines { get; }
    /// <summary>Plain-text heading text for compatibility, slugs, and TOC labels.</summary>
    public string Text { get; }
    /// <summary>
    /// Creates a new heading block.
    /// </summary>
    /// <param name="level">Desired level; constrained to [1,6].</param>
    /// <param name="text">Heading text.</param>
    public HeadingBlock(int level, string text)
        : this(level, CreateTextInlines(text)) {
    }

    /// <summary>
    /// Creates a new heading block from parsed inline content.
    /// </summary>
    /// <param name="level">Desired level; constrained to [1,6].</param>
    /// <param name="inlines">Inline content.</param>
    public HeadingBlock(int level, InlineSequence inlines) {
        // Manual clamp to support netstandard2.0 where Math.Clamp may not exist.
        if (level < 1) level = 1; else if (level > 6) level = 6;
        Level = level;
        Inlines = inlines ?? new InlineSequence();
        Text = InlinePlainText.Extract(Inlines);
    }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => new string('#', Level) + " " + Inlines.RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var id = MarkdownSlug.GitHub(Text);
        return $"<h{Level} id=\"{id}\">{Inlines.RenderHtml()}</h{Level}>";
    }

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        if (!context.HeadingSlugs.TryGetValue(this, out var id)) {
            id = MarkdownSlug.GitHub(Text);
        }

        var sb = new System.Text.StringBuilder();
        sb.Append("<h").Append(Level).Append(" id=\"").Append(id).Append("\">");
        sb.Append(Inlines.RenderHtml());
        if (context.Options.IncludeAnchorLinks || context.Options.ShowAnchorIcons) {
            var icon = System.Net.WebUtility.HtmlEncode(context.Options.AnchorIcon ?? "🔗");
            sb.Append("<a class=\"heading-anchor\" href=\"#")
              .Append(id)
              .Append("\" data-anchor-id=\"")
              .Append(id)
              .Append("\" title=\"Copy link\" aria-label=\"Copy link\">")
              .Append(icon)
              .Append("</a>");
        }
        sb.Append("</h").Append(Level).Append('>');

        if (context.Options.BackToTopLinks && Level >= context.Options.BackToTopMinLevel) {
            var text = System.Net.WebUtility.HtmlEncode(context.Options.BackToTopText ?? "Back to top");
            sb.Append("<div class=\"back-to-top\"><a href=\"#top\">").Append(text).Append("</a></div>");
        }

        return sb.ToString();
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildHeadingBlock(this, span);

    private static InlineSequence CreateTextInlines(string? text) {
        var inlines = new InlineSequence();
        if (!string.IsNullOrEmpty(text)) {
            inlines.Text(text!);
        }
        return inlines;
    }
}
