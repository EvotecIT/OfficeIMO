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
    internal bool HasOpeningMarkerSourceInfo { get; private set; }
    internal int OpeningMarkerSourceLineOffset { get; private set; }
    internal int OpeningMarkerSourceStartColumn { get; private set; }
    internal int OpeningMarkerSourceEndColumn { get; private set; }
    internal bool HasSetextUnderlineMarkerSourceInfo { get; private set; }
    internal int SetextUnderlineMarkerSourceLineOffset { get; private set; }
    internal int SetextUnderlineMarkerSourceStartColumn { get; private set; }
    internal int SetextUnderlineMarkerSourceEndColumn { get; private set; }
    internal bool HasTextSourceInfo { get; private set; }
    internal int TextSourceLineOffset { get; private set; }
    internal int TextSourceEndLineOffset { get; private set; }
    internal int TextSourceStartColumn { get; private set; }
    internal int TextSourceEndColumn { get; private set; }
    internal bool HasClosingMarkerSourceInfo { get; private set; }
    internal int ClosingMarkerSourceLineOffset { get; private set; }
    internal int ClosingMarkerSourceStartColumn { get; private set; }
    internal int ClosingMarkerSourceEndColumn { get; private set; }
    /// <summary>Source span for the ATX opening marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; private set; }
    /// <summary>Exact ATX opening marker token when parsed from markdown.</summary>
    public string? OpeningMarkerText { get; private set; }
    /// <summary>Source span for a Setext underline marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? SetextUnderlineMarkerSourceSpan { get; private set; }
    /// <summary>Exact Setext underline marker token when parsed from markdown.</summary>
    public string? SetextUnderlineMarkerText { get; private set; }
    /// <summary>Source span for an optional ATX closing marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? ClosingMarkerSourceSpan { get; private set; }
    /// <summary>Exact optional ATX closing marker token when parsed from markdown.</summary>
    public string? ClosingMarkerText { get; private set; }
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

    internal void SetOpeningMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, MarkdownSourceSpan? sourceSpan = null) {
        HasOpeningMarkerSourceInfo = true;
        OpeningMarkerSourceLineOffset = Math.Max(0, lineOffset);
        OpeningMarkerSourceStartColumn = Math.Max(1, startColumn);
        OpeningMarkerSourceEndColumn = Math.Max(OpeningMarkerSourceStartColumn, endColumn);
        OpeningMarkerSourceSpan = sourceSpan;
        OpeningMarkerText = new string('#', OpeningMarkerSourceEndColumn - OpeningMarkerSourceStartColumn + 1);
    }

    internal void SetSetextUnderlineMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, string markerText, MarkdownSourceSpan? sourceSpan = null) {
        HasSetextUnderlineMarkerSourceInfo = true;
        SetextUnderlineMarkerSourceLineOffset = Math.Max(0, lineOffset);
        SetextUnderlineMarkerSourceStartColumn = Math.Max(1, startColumn);
        SetextUnderlineMarkerSourceEndColumn = Math.Max(SetextUnderlineMarkerSourceStartColumn, endColumn);
        SetextUnderlineMarkerSourceSpan = sourceSpan;
        SetextUnderlineMarkerText = markerText ?? string.Empty;
    }

    internal void SetTextSourceInfo(int lineOffset, int startColumn, int endColumn) {
        SetTextSourceInfo(lineOffset, startColumn, lineOffset, endColumn);
    }

    internal void SetTextSourceInfo(int startLineOffset, int startColumn, int endLineOffset, int endColumn) {
        HasTextSourceInfo = true;
        TextSourceLineOffset = Math.Max(0, startLineOffset);
        TextSourceEndLineOffset = Math.Max(TextSourceLineOffset, endLineOffset);
        TextSourceStartColumn = Math.Max(1, startColumn);
        TextSourceEndColumn = Math.Max(1, endColumn);
        if (TextSourceEndLineOffset == TextSourceLineOffset) {
            TextSourceEndColumn = Math.Max(TextSourceStartColumn, TextSourceEndColumn);
        }
    }

    internal void SetClosingMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, MarkdownSourceSpan? sourceSpan = null) {
        HasClosingMarkerSourceInfo = true;
        ClosingMarkerSourceLineOffset = Math.Max(0, lineOffset);
        ClosingMarkerSourceStartColumn = Math.Max(1, startColumn);
        ClosingMarkerSourceEndColumn = Math.Max(ClosingMarkerSourceStartColumn, endColumn);
        ClosingMarkerSourceSpan = sourceSpan;
        ClosingMarkerText = new string('#', ClosingMarkerSourceEndColumn - ClosingMarkerSourceStartColumn + 1);
    }

    internal void OffsetRelativeSourceInfoLines(int lineOffsetDelta) {
        if (lineOffsetDelta <= 0) {
            return;
        }

        if (HasLevelSourceInfo) {
            LevelSourceLineOffset += lineOffsetDelta;
        }

        if (HasOpeningMarkerSourceInfo) {
            OpeningMarkerSourceLineOffset += lineOffsetDelta;
        }

        if (HasSetextUnderlineMarkerSourceInfo) {
            SetextUnderlineMarkerSourceLineOffset += lineOffsetDelta;
        }

        if (HasTextSourceInfo) {
            TextSourceLineOffset += lineOffsetDelta;
            TextSourceEndLineOffset += lineOffsetDelta;
        }

        if (HasClosingMarkerSourceInfo) {
            ClosingMarkerSourceLineOffset += lineOffsetDelta;
        }
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => new string('#', Level) + " " + Inlines.RenderMarkdown() + MarkdownAttributeBlockRenderer.RenderTrailing(Attributes);
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var id = MarkdownSlug.Generate(Text, MarkdownHeadingIdentifierStyle.OfficeIMO);
        return $"<h{Level}{MarkdownHtmlAttributes.Render(Attributes, null, id)}>{Inlines.RenderHtml()}</h{Level}>";
    }

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        var id = context.Options.AutoHeadingIdentifiers
            ? context.HeadingCatalog.GetHeadingAnchor(this)
            : string.Empty;

        var sb = new System.Text.StringBuilder();
        sb.Append("<h").Append(Level);
        var effectiveId = !string.IsNullOrWhiteSpace(Attributes.ElementId) ? Attributes.ElementId : id;
        sb.Append(MarkdownHtmlAttributes.Render(Attributes, context.Options, id));
        sb.Append(">");
        sb.Append(Inlines.RenderHtml());
        if (!string.IsNullOrEmpty(effectiveId) && (context.Options.IncludeAnchorLinks || context.Options.ShowAnchorIcons)) {
            var icon = HtmlTextEncoder.Encode(context.Options.AnchorIcon ?? "🔗", context.Options);
            sb.Append("<a class=\"heading-anchor\" href=\"#")
              .Append(HtmlTextEncoder.Encode(effectiveId, context.Options))
              .Append("\" data-anchor-id=\"")
              .Append(HtmlTextEncoder.Encode(effectiveId, context.Options))
              .Append("\" title=\"Copy link\" aria-label=\"Copy link\">")
              .Append(icon)
              .Append("</a>");
        }
        sb.Append("</h").Append(Level).Append('>');

        if (context.Options.BackToTopLinks && Level >= context.Options.BackToTopMinLevel) {
            var text = HtmlTextEncoder.Encode(context.Options.BackToTopText ?? "Back to top", context.Options);
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

        var openingMarkerSpan = GetOpeningMarkerSourceSpan(span);
        if (openingMarkerSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HeadingOpeningMarker,
                openingMarkerSpan,
                OpeningMarkerText ?? new string('#', Level)));
        }

        var setextUnderlineMarkerSpan = GetSetextUnderlineMarkerSourceSpan(span);
        if (setextUnderlineMarkerSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HeadingSetextUnderlineMarker,
                setextUnderlineMarkerSpan,
                SetextUnderlineMarkerText));
        }

        var closingMarkerSpan = GetClosingMarkerSourceSpan(span);
        if (closingMarkerSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HeadingClosingMarker,
                closingMarkerSpan,
                ClosingMarkerText ?? "#"));
        }

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
        if (OpeningMarkerSourceSpan.HasValue) {
            return OpeningMarkerSourceSpan;
        }

        if (SetextUnderlineMarkerSourceSpan.HasValue) {
            return SetextUnderlineMarkerSourceSpan;
        }

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
        if (Inlines.SourceSpan.HasValue) {
            return Inlines.SourceSpan;
        }

        if (!span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        if (HasTextSourceInfo) {
            return new MarkdownSourceSpan(
                value.StartLine + TextSourceLineOffset,
                TextSourceStartColumn,
                value.StartLine + TextSourceEndLineOffset,
                TextSourceEndColumn);
        }

        return Inlines.SourceSpan;
    }

    private MarkdownSourceSpan? GetOpeningMarkerSourceSpan(MarkdownSourceSpan? span) {
        if (OpeningMarkerSourceSpan.HasValue) {
            return OpeningMarkerSourceSpan;
        }

        if (!HasOpeningMarkerSourceInfo || !span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        OpeningMarkerSourceSpan = new MarkdownSourceSpan(
            value.StartLine + OpeningMarkerSourceLineOffset,
            OpeningMarkerSourceStartColumn,
            value.StartLine + OpeningMarkerSourceLineOffset,
            OpeningMarkerSourceEndColumn);
        return OpeningMarkerSourceSpan;
    }

    private MarkdownSourceSpan? GetSetextUnderlineMarkerSourceSpan(MarkdownSourceSpan? span) {
        if (SetextUnderlineMarkerSourceSpan.HasValue) {
            return SetextUnderlineMarkerSourceSpan;
        }

        if (!HasSetextUnderlineMarkerSourceInfo || !span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        SetextUnderlineMarkerSourceSpan = new MarkdownSourceSpan(
            value.StartLine + SetextUnderlineMarkerSourceLineOffset,
            SetextUnderlineMarkerSourceStartColumn,
            value.StartLine + SetextUnderlineMarkerSourceLineOffset,
            SetextUnderlineMarkerSourceEndColumn);
        return SetextUnderlineMarkerSourceSpan;
    }

    private MarkdownSourceSpan? GetClosingMarkerSourceSpan(MarkdownSourceSpan? span) {
        if (ClosingMarkerSourceSpan.HasValue) {
            return ClosingMarkerSourceSpan;
        }

        if (!HasClosingMarkerSourceInfo || !span.HasValue || !span.Value.StartColumn.HasValue) {
            return null;
        }

        var value = span.Value;
        ClosingMarkerSourceSpan = new MarkdownSourceSpan(
            value.StartLine + ClosingMarkerSourceLineOffset,
            ClosingMarkerSourceStartColumn,
            value.StartLine + ClosingMarkerSourceLineOffset,
            ClosingMarkerSourceEndColumn);
        return ClosingMarkerSourceSpan;
    }
}
