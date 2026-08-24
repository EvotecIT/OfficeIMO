namespace OfficeIMO.Markdown;

/// <summary>
/// Markdown heading (ATX) block, levels 1–6.
/// </summary>
public sealed class HeadingBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock, IContextualHtmlMarkdownBlock, IHeadingMarkdownBlock {
    private HeadingBlockSyntaxMetadata? _syntaxMetadata;

    /// <summary>Heading level constrained to [1,6].</summary>
    public int Level { get; }
    /// <summary>Inline content owned by this heading.</summary>
    public InlineSequence Inlines { get; }
    /// <summary>Plain-text heading text for compatibility, slugs, and TOC labels.</summary>
    public string Text { get; }
    internal bool HasLevelSourceInfo => _syntaxMetadata?.HasLevelSourceInfo == true;
    internal int LevelSourceLineOffset => _syntaxMetadata?.LevelSourceLineOffset ?? 0;
    internal int LevelSourceStartColumn => _syntaxMetadata?.LevelSourceStartColumn ?? 0;
    internal int LevelSourceEndColumn => _syntaxMetadata?.LevelSourceEndColumn ?? 0;
    internal bool HasOpeningMarkerSourceInfo => _syntaxMetadata?.HasOpeningMarkerSourceInfo == true;
    internal int OpeningMarkerSourceLineOffset => _syntaxMetadata?.OpeningMarkerSourceLineOffset ?? 0;
    internal int OpeningMarkerSourceStartColumn => _syntaxMetadata?.OpeningMarkerSourceStartColumn ?? 0;
    internal int OpeningMarkerSourceEndColumn => _syntaxMetadata?.OpeningMarkerSourceEndColumn ?? 0;
    internal bool HasSetextUnderlineMarkerSourceInfo => _syntaxMetadata?.HasSetextUnderlineMarkerSourceInfo == true;
    internal int SetextUnderlineMarkerSourceLineOffset => _syntaxMetadata?.SetextUnderlineMarkerSourceLineOffset ?? 0;
    internal int SetextUnderlineMarkerSourceStartColumn => _syntaxMetadata?.SetextUnderlineMarkerSourceStartColumn ?? 0;
    internal int SetextUnderlineMarkerSourceEndColumn => _syntaxMetadata?.SetextUnderlineMarkerSourceEndColumn ?? 0;
    internal bool HasTextSourceInfo => _syntaxMetadata?.HasTextSourceInfo == true;
    internal int TextSourceLineOffset => _syntaxMetadata?.TextSourceLineOffset ?? 0;
    internal int TextSourceEndLineOffset => _syntaxMetadata?.TextSourceEndLineOffset ?? 0;
    internal int TextSourceStartColumn => _syntaxMetadata?.TextSourceStartColumn ?? 0;
    internal int TextSourceEndColumn => _syntaxMetadata?.TextSourceEndColumn ?? 0;
    internal bool HasClosingMarkerSourceInfo => _syntaxMetadata?.HasClosingMarkerSourceInfo == true;
    internal int ClosingMarkerSourceLineOffset => _syntaxMetadata?.ClosingMarkerSourceLineOffset ?? 0;
    internal int ClosingMarkerSourceStartColumn => _syntaxMetadata?.ClosingMarkerSourceStartColumn ?? 0;
    internal int ClosingMarkerSourceEndColumn => _syntaxMetadata?.ClosingMarkerSourceEndColumn ?? 0;
    internal bool SuppressAutoIdentifier { get; private set; }
    /// <summary>Source span for the heading marker or setext underline that determines the level.</summary>
    public MarkdownSourceSpan? LevelSourceSpan {
        get => _syntaxMetadata?.LevelSourceSpan;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.LevelSourceSpan = current);
    }
    /// <summary>Source span for the heading text payload.</summary>
    public MarkdownSourceSpan? TextSourceSpan {
        get => _syntaxMetadata?.TextSourceSpan;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.TextSourceSpan = current);
    }
    /// <summary>Source span for the ATX opening marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan {
        get => _syntaxMetadata?.OpeningMarkerSourceSpan;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.OpeningMarkerSourceSpan = current);
    }
    /// <summary>Exact ATX opening marker token when parsed from markdown.</summary>
    public string? OpeningMarkerText {
        get => _syntaxMetadata?.OpeningMarkerText;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.OpeningMarkerText = current);
    }
    /// <summary>Source span for a Setext underline marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? SetextUnderlineMarkerSourceSpan {
        get => _syntaxMetadata?.SetextUnderlineMarkerSourceSpan;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.SetextUnderlineMarkerSourceSpan = current);
    }
    /// <summary>Exact Setext underline marker token when parsed from markdown.</summary>
    public string? SetextUnderlineMarkerText {
        get => _syntaxMetadata?.SetextUnderlineMarkerText;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.SetextUnderlineMarkerText = current);
    }
    /// <summary>Source span for an optional ATX closing marker token when parsed from markdown.</summary>
    public MarkdownSourceSpan? ClosingMarkerSourceSpan {
        get => _syntaxMetadata?.ClosingMarkerSourceSpan;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.ClosingMarkerSourceSpan = current);
    }
    /// <summary>Exact optional ATX closing marker token when parsed from markdown.</summary>
    public string? ClosingMarkerText {
        get => _syntaxMetadata?.ClosingMarkerText;
        private set => SetSyntaxValue(value, static (metadata, current) => metadata.ClosingMarkerText = current);
    }
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
        Text = Inlines.Nodes.Count == 1 && Inlines.Nodes[0] is MarkdownTextRun textRun
            ? textRun.Text
            : InlinePlainText.Extract(Inlines);
        TextSourceSpan = Inlines.SourceSpan;
    }

    internal void SetLevelSourceInfo(int lineOffset, int startColumn, int endColumn) {
        var metadata = SyntaxMetadata;
        metadata.HasLevelSourceInfo = true;
        metadata.LevelSourceLineOffset = Math.Max(0, lineOffset);
        metadata.LevelSourceStartColumn = Math.Max(1, startColumn);
        metadata.LevelSourceEndColumn = Math.Max(metadata.LevelSourceStartColumn, endColumn);
    }

    internal void SetOpeningMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, MarkdownSourceSpan? sourceSpan = null) {
        var metadata = SyntaxMetadata;
        metadata.HasOpeningMarkerSourceInfo = true;
        metadata.OpeningMarkerSourceLineOffset = Math.Max(0, lineOffset);
        metadata.OpeningMarkerSourceStartColumn = Math.Max(1, startColumn);
        metadata.OpeningMarkerSourceEndColumn = Math.Max(metadata.OpeningMarkerSourceStartColumn, endColumn);
        OpeningMarkerSourceSpan = sourceSpan;
        LevelSourceSpan = sourceSpan ?? LevelSourceSpan;
        OpeningMarkerText = new string('#', metadata.OpeningMarkerSourceEndColumn - metadata.OpeningMarkerSourceStartColumn + 1);
    }

    internal void SetSetextUnderlineMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, string markerText, MarkdownSourceSpan? sourceSpan = null) {
        var metadata = SyntaxMetadata;
        metadata.HasSetextUnderlineMarkerSourceInfo = true;
        metadata.SetextUnderlineMarkerSourceLineOffset = Math.Max(0, lineOffset);
        metadata.SetextUnderlineMarkerSourceStartColumn = Math.Max(1, startColumn);
        metadata.SetextUnderlineMarkerSourceEndColumn = Math.Max(metadata.SetextUnderlineMarkerSourceStartColumn, endColumn);
        SetextUnderlineMarkerSourceSpan = sourceSpan;
        LevelSourceSpan = sourceSpan ?? LevelSourceSpan;
        SetextUnderlineMarkerText = markerText ?? string.Empty;
    }

    internal void SetTextSourceInfo(int lineOffset, int startColumn, int endColumn) {
        SetTextSourceInfo(lineOffset, startColumn, lineOffset, endColumn);
    }

    internal void SetTextSourceInfo(int startLineOffset, int startColumn, int endLineOffset, int endColumn) {
        var metadata = SyntaxMetadata;
        metadata.HasTextSourceInfo = true;
        metadata.TextSourceLineOffset = Math.Max(0, startLineOffset);
        metadata.TextSourceEndLineOffset = Math.Max(metadata.TextSourceLineOffset, endLineOffset);
        metadata.TextSourceStartColumn = Math.Max(1, startColumn);
        metadata.TextSourceEndColumn = Math.Max(1, endColumn);
        if (metadata.TextSourceEndLineOffset == metadata.TextSourceLineOffset) {
            metadata.TextSourceEndColumn = Math.Max(metadata.TextSourceStartColumn, metadata.TextSourceEndColumn);
        }

        TextSourceSpan = Inlines.SourceSpan ?? TextSourceSpan;
    }

    internal void SetClosingMarkerSourceInfo(int lineOffset, int startColumn, int endColumn, MarkdownSourceSpan? sourceSpan = null) {
        var metadata = SyntaxMetadata;
        metadata.HasClosingMarkerSourceInfo = true;
        metadata.ClosingMarkerSourceLineOffset = Math.Max(0, lineOffset);
        metadata.ClosingMarkerSourceStartColumn = Math.Max(1, startColumn);
        metadata.ClosingMarkerSourceEndColumn = Math.Max(metadata.ClosingMarkerSourceStartColumn, endColumn);
        ClosingMarkerSourceSpan = sourceSpan;
        ClosingMarkerText = new string('#', metadata.ClosingMarkerSourceEndColumn - metadata.ClosingMarkerSourceStartColumn + 1);
    }

    internal void SuppressAutomaticIdentifier() {
        SuppressAutoIdentifier = true;
    }

    internal void OffsetRelativeSourceInfoLines(int lineOffsetDelta) {
        if (lineOffsetDelta <= 0 || _syntaxMetadata == null) {
            return;
        }

        if (_syntaxMetadata.HasLevelSourceInfo) {
            _syntaxMetadata.LevelSourceLineOffset += lineOffsetDelta;
        }

        if (_syntaxMetadata.HasOpeningMarkerSourceInfo) {
            _syntaxMetadata.OpeningMarkerSourceLineOffset += lineOffsetDelta;
        }

        if (_syntaxMetadata.HasSetextUnderlineMarkerSourceInfo) {
            _syntaxMetadata.SetextUnderlineMarkerSourceLineOffset += lineOffsetDelta;
        }

        if (_syntaxMetadata.HasTextSourceInfo) {
            _syntaxMetadata.TextSourceLineOffset += lineOffsetDelta;
            _syntaxMetadata.TextSourceEndLineOffset += lineOffsetDelta;
        }

        if (_syntaxMetadata.HasClosingMarkerSourceInfo) {
            _syntaxMetadata.ClosingMarkerSourceLineOffset += lineOffsetDelta;
        }
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => new string('#', Level) + " " + Inlines.RenderMarkdown() + MarkdownAttributeBlockRenderer.RenderTrailing(Attributes);
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var id = SuppressAutoIdentifier ? string.Empty : MarkdownSlug.Generate(Text, MarkdownHeadingIdentifierStyle.OfficeIMO);
        return $"<h{Level}{MarkdownHtmlAttributes.Render(Attributes, null, id)}>{Inlines.RenderHtml()}</h{Level}>";
    }

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        var id = !SuppressAutoIdentifier && context.Options.AutoHeadingIdentifiers
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
                ResolveLevelSourceSpan(span),
                literal: Level.ToString(System.Globalization.CultureInfo.InvariantCulture))
        };

        nodes.Add(MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(
            MarkdownSyntaxKind.HeadingText,
            Inlines,
            ResolveTextSourceSpan(span),
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

    private MarkdownSourceSpan? ResolveLevelSourceSpan(MarkdownSourceSpan? span) {
        if (OpeningMarkerSourceSpan.HasValue) {
            LevelSourceSpan = OpeningMarkerSourceSpan;
            return LevelSourceSpan;
        }

        if (SetextUnderlineMarkerSourceSpan.HasValue) {
            LevelSourceSpan = SetextUnderlineMarkerSourceSpan;
            return LevelSourceSpan;
        }

        if (!span.HasValue || !span.Value.StartColumn.HasValue) {
            return LevelSourceSpan;
        }

        var value = span.Value;
        if (HasLevelSourceInfo) {
            LevelSourceSpan = new MarkdownSourceSpan(
                value.StartLine + LevelSourceLineOffset,
                LevelSourceStartColumn,
                value.StartLine + LevelSourceLineOffset,
                LevelSourceEndColumn);
            return LevelSourceSpan;
        }

        if (value.EndLine > value.StartLine && value.EndColumn.HasValue) {
            LevelSourceSpan = new MarkdownSourceSpan(value.EndLine, 1, value.EndLine, value.EndColumn.Value);
            return LevelSourceSpan;
        }

        var startColumn = value.StartColumn.Value;
        LevelSourceSpan = new MarkdownSourceSpan(value.StartLine, startColumn, value.StartLine, startColumn + Level - 1);
        return LevelSourceSpan;
    }

    private MarkdownSourceSpan? ResolveTextSourceSpan(MarkdownSourceSpan? span) {
        if (Inlines.SourceSpan.HasValue) {
            TextSourceSpan = Inlines.SourceSpan;
            return TextSourceSpan;
        }

        if (!span.HasValue || !span.Value.StartColumn.HasValue) {
            return TextSourceSpan;
        }

        var value = span.Value;
        if (HasTextSourceInfo) {
            TextSourceSpan = new MarkdownSourceSpan(
                value.StartLine + TextSourceLineOffset,
                TextSourceStartColumn,
                value.StartLine + TextSourceEndLineOffset,
                TextSourceEndColumn);
            return TextSourceSpan;
        }

        return TextSourceSpan;
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

    private HeadingBlockSyntaxMetadata SyntaxMetadata =>
        _syntaxMetadata ??= new HeadingBlockSyntaxMetadata();

    private void SetSyntaxValue(
        MarkdownSourceSpan? value,
        Action<HeadingBlockSyntaxMetadata, MarkdownSourceSpan?> setter) {
        if (value.HasValue) {
            setter(SyntaxMetadata, value);
        } else if (_syntaxMetadata != null) {
            setter(_syntaxMetadata, null);
        }
    }

    private void SetSyntaxValue(
        string? value,
        Action<HeadingBlockSyntaxMetadata, string?> setter) {
        if (value != null) {
            setter(SyntaxMetadata, value);
        } else if (_syntaxMetadata != null) {
            setter(_syntaxMetadata, null);
        }
    }
}

internal sealed class HeadingBlockSyntaxMetadata {
    internal bool HasLevelSourceInfo;
    internal int LevelSourceLineOffset;
    internal int LevelSourceStartColumn;
    internal int LevelSourceEndColumn;
    internal bool HasOpeningMarkerSourceInfo;
    internal int OpeningMarkerSourceLineOffset;
    internal int OpeningMarkerSourceStartColumn;
    internal int OpeningMarkerSourceEndColumn;
    internal bool HasSetextUnderlineMarkerSourceInfo;
    internal int SetextUnderlineMarkerSourceLineOffset;
    internal int SetextUnderlineMarkerSourceStartColumn;
    internal int SetextUnderlineMarkerSourceEndColumn;
    internal bool HasTextSourceInfo;
    internal int TextSourceLineOffset;
    internal int TextSourceEndLineOffset;
    internal int TextSourceStartColumn;
    internal int TextSourceEndColumn;
    internal bool HasClosingMarkerSourceInfo;
    internal int ClosingMarkerSourceLineOffset;
    internal int ClosingMarkerSourceStartColumn;
    internal int ClosingMarkerSourceEndColumn;
    internal MarkdownSourceSpan? LevelSourceSpan;
    internal MarkdownSourceSpan? TextSourceSpan;
    internal MarkdownSourceSpan? OpeningMarkerSourceSpan;
    internal string? OpeningMarkerText;
    internal MarkdownSourceSpan? SetextUnderlineMarkerSourceSpan;
    internal string? SetextUnderlineMarkerText;
    internal MarkdownSourceSpan? ClosingMarkerSourceSpan;
    internal string? ClosingMarkerText;
}
