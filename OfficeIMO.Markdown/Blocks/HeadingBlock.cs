namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown heading (ATX) block, levels 1–6.
/// </summary>
public sealed class HeadingBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IContextualHtmlMarkdownBlock, IHeadingMarkdownBlock {
    /// <summary>Heading level constrained to [1,6].</summary>
    public int Level { get; }
    /// <summary>Inline content owned by this heading.</summary>
    public InlineSequence Inlines { get; }
    /// <summary>Plain-text heading text for compatibility, slugs, and TOC labels.</summary>
    public string Text { get; }
    internal bool HasLevelSourceInfo { get; private set; }
    internal int LevelSourceLineOffset { get; private set; }
    internal int LevelSourceStartColumn { get; private set; }
    internal int LevelSourceEndColumn { get; private set; }
    internal bool HasTextSourceInfo { get; private set; }
    internal int TextSourceLineOffset { get; private set; }
    internal int TextSourceStartColumn { get; private set; }
    internal int TextSourceEndColumn { get; private set; }
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

    internal void SetLevelSourceInfo(int lineOffset, int startColumn, int endColumn) {
        HasLevelSourceInfo = true;
        LevelSourceLineOffset = Math.Max(0, lineOffset);
        LevelSourceStartColumn = Math.Max(1, startColumn);
        LevelSourceEndColumn = Math.Max(LevelSourceStartColumn, endColumn);
    }

    internal void SetTextSourceInfo(int lineOffset, int startColumn, int endColumn) {
        HasTextSourceInfo = true;
        TextSourceLineOffset = Math.Max(0, lineOffset);
        TextSourceStartColumn = Math.Max(1, startColumn);
        TextSourceEndColumn = Math.Max(TextSourceStartColumn, endColumn);
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => new string('#', Level) + " " + Inlines.RenderMarkdown();
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var id = MarkdownSlug.GitHub(Text);
        return $"<h{Level} id=\"{id}\">{Inlines.RenderHtml()}</h{Level}>";
    }

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        var id = context.HeadingCatalog.GetHeadingAnchor(this);

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

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode> {
            new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HeadingLevel,
                GetLevelSourceSpan(span),
                literal: Level.ToString(System.Globalization.CultureInfo.InvariantCulture))
        };

        nodes.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
            MarkdownSyntaxKind.HeadingText,
            Inlines,
            GetTextSourceSpan(span),
            Inlines.RenderMarkdown()));

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.Heading, span, Inlines.RenderMarkdown(), nodes, this);
    }

    private static InlineSequence CreateTextInlines(string? text) {
        var inlines = new InlineSequence();
        if (!string.IsNullOrEmpty(text)) {
            inlines.Text(text!);
        }
        return inlines;
    }

    private MarkdownSourceSpan? GetLevelSourceSpan(MarkdownSourceSpan? span) {
        if (!span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        if (HasLevelSourceInfo) {
            return new MarkdownSourceSpan(
                value.StartLine + LevelSourceLineOffset,
                LevelSourceStartColumn,
                value.StartLine + LevelSourceLineOffset,
                LevelSourceEndColumn);
        }

        if (value.EndLine > value.StartLine && value.EndColumn.HasValue) {
            return new MarkdownSourceSpan(value.EndLine, 1, value.EndLine, value.EndColumn.Value);
        }

        var startColumn = value.StartColumn.Value;
        return new MarkdownSourceSpan(value.StartLine, startColumn, value.StartLine, startColumn + Level - 1);
    }

    private MarkdownSourceSpan? GetTextSourceSpan(MarkdownSourceSpan? span) {
        if (!span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        if (HasTextSourceInfo) {
            return new MarkdownSourceSpan(
                value.StartLine + TextSourceLineOffset,
                TextSourceStartColumn,
                value.StartLine + TextSourceLineOffset,
                TextSourceEndColumn);
        }

        return Inlines.SourceSpan;
    }
}
