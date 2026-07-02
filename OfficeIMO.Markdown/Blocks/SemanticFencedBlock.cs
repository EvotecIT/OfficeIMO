namespace OfficeIMO.Markdown;

/// <summary>
/// First-class AST node for fenced blocks whose language maps to host-defined semantics
/// such as diagrams, charts, data views, or other non-code contracts.
/// </summary>
public sealed class SemanticFencedBlock : MarkdownBlock, IMarkdownBlock, ICaptionable, ISyntaxMarkdownBlock {
    /// <summary>Create a semantic fenced block.</summary>
    public SemanticFencedBlock(string semanticKind, string language, string content, string? caption = null)
        : this(semanticKind, language, content, caption, isFenced: true) {
    }

    internal SemanticFencedBlock(string semanticKind, string language, string content, string? caption, bool isFenced) {
        SemanticKind = string.IsNullOrWhiteSpace(semanticKind) ? MarkdownSemanticKinds.Custom : semanticKind.Trim();
        FenceInfo = MarkdownCodeFenceInfo.Parse(language);
        InfoString = FenceInfo.InfoString;
        Language = FenceInfo.Language;
        Content = NormalizeLineEndings(content);
        Caption = caption;
        IsFenced = isFenced;
        SetAttributes(MarkdownAttributeSet.Create(FenceInfo.ElementId, FenceInfo.Classes, FenceInfo.Attributes));
    }

    /// <summary>Host-defined semantic contract for this block (for example <c>chart</c> or <c>mermaid</c>).</summary>
    public string SemanticKind { get; }

    /// <summary>Original fence language / info string.</summary>
    public string Language { get; }

    /// <summary>Full original fence info string.</summary>
    public string InfoString { get; }

    /// <summary>Structured fenced-code info metadata.</summary>
    public MarkdownCodeFenceInfo FenceInfo { get; }

    /// <summary>Raw fenced payload.</summary>
    public string Content { get; }

    /// <summary>Optional caption shown under the block.</summary>
    public string? Caption { get; set; }

    internal bool IsFenced { get; }
    internal int FenceIndentColumns { get; private set; }
    internal char FenceChar { get; private set; } = '`';
    internal int FenceLength { get; private set; } = 3;
    internal int FenceInfoPaddingColumns { get; private set; }
    internal int FenceInfoPaddingCharacters { get; private set; }
    internal bool HasClosingFence { get; private set; } = true;
    internal int ClosingFenceIndentColumns { get; private set; }
    internal int ClosingFenceLength { get; private set; } = 3;
    internal int? FencedContentLineCount { get; private set; }
    /// <summary>Source span for the opening fence token when parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; internal set; }
    /// <summary>Source span for the fenced-block info string when parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? InfoStringSourceSpan { get; internal set; }
    /// <summary>Source span for the fenced payload when source-backed.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; internal set; }
    /// <summary>Source span for the closing fence token when parsed from a closed fenced source block.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; internal set; }

    internal void SetFenceSourceInfo(
        int fenceIndentColumns,
        int fenceLength,
        int infoPaddingColumns,
        int infoPaddingCharacters,
        char fenceChar = '`',
        bool hasClosingFence = true,
        int closingFenceIndentColumns = 0,
        int closingFenceLength = 3,
        int? fencedContentLineCount = null) {
        FenceIndentColumns = Math.Max(0, fenceIndentColumns);
        FenceChar = fenceChar == '~' ? '~' : '`';
        FenceLength = Math.Max(3, fenceLength);
        FenceInfoPaddingColumns = Math.Max(0, infoPaddingColumns);
        FenceInfoPaddingCharacters = Math.Max(0, infoPaddingCharacters);
        HasClosingFence = hasClosingFence;
        ClosingFenceIndentColumns = Math.Max(0, closingFenceIndentColumns);
        ClosingFenceLength = Math.Max(3, closingFenceLength);
        FencedContentLineCount = fencedContentLineCount.HasValue && fencedContentLineCount.Value >= 0
            ? fencedContentLineCount.Value
            : null;
    }

    internal void SetFenceTokenSourceSpans(
        MarkdownSourceSpan? openingFenceSourceSpan,
        MarkdownSourceSpan? infoStringSourceSpan,
        MarkdownSourceSpan? contentSourceSpan,
        MarkdownSourceSpan? closingFenceSourceSpan) {
        OpeningFenceSourceSpan = openingFenceSourceSpan;
        InfoStringSourceSpan = infoStringSourceSpan;
        ContentSourceSpan = contentSourceSpan;
        ClosingFenceSourceSpan = closingFenceSourceSpan;
    }

    string IMarkdownBlock.RenderMarkdown() {
        string fence = MarkdownFence.BuildSafeFence(Content);

        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{fence}{InfoString}");
        sb.AppendLine(Content);
        sb.AppendLine(fence);
        if (!string.IsNullOrWhiteSpace(Caption)) {
            sb.AppendLine("_" + Caption + "_");
        }

        return sb.ToString().TrimEnd();
    }

    string IMarkdownBlock.RenderHtml() {
        var options = HtmlRenderContext.Options;
        var overridden = options?.SemanticFencedBlockHtmlRenderer?.Invoke(this, options);
        if (overridden != null) {
            return overridden;
        }

        var fallbackBlock = new CodeBlock(InfoString, Content) {
            Caption = Caption
        };
        fallbackBlock.SetAttributes(Attributes);

        var codeFallback = options?.CodeBlockHtmlRenderer?.Invoke(fallbackBlock, options);
        if (codeFallback != null) {
            return codeFallback;
        }

        string attrs = MarkdownHtmlAttributes.Render(Attributes, options);
        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{HtmlTextEncoder.Encode(Language, options)}\"";
        string code = HtmlTextEncoder.Encode(Content, options);
        if (code.Length > 0) {
            code += "\n";
        }

        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{HtmlTextEncoder.Encode(Caption!, options)}</div>";
        return $"<pre{attrs}><code{lang}>{code}</code></pre>{caption}";
    }

    private static string NormalizeLineEndings(string? content) {
        if (string.IsNullOrEmpty(content)) {
            return string.Empty;
        }

        return content!
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode> {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.FenceSemanticKind, literal: SemanticKind)
        };

        OpeningFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetOpeningFenceSpan(span, IsFenced, FenceIndentColumns, FenceLength);
        ClosingFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetClosingFenceSpan(span, IsFenced, Content, HasClosingFence, ClosingFenceIndentColumns, ClosingFenceLength, FencedContentLineCount);

        if (OpeningFenceSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceOpening,
                OpeningFenceSourceSpan,
                new string(FenceChar, FenceLength)));
        }

        InfoStringSourceSpan = MarkdownFencedBlockSourceSpans.GetInfoSpan(span, IsFenced, InfoString, FenceIndentColumns, FenceLength, FenceInfoPaddingColumns, FenceInfoPaddingCharacters);
        if (InfoStringSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                InfoStringSourceSpan,
                InfoString));
        }

        ContentSourceSpan = MarkdownFencedBlockSourceSpans.GetContentSpan(span, IsFenced, Content, FencedContentLineCount);
        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeContent,
            ContentSourceSpan,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content)));

        if (ClosingFenceSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceClosing,
                ClosingFenceSourceSpan,
                new string(FenceChar, ClosingFenceLength)));
        }

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.SemanticFencedBlock,
            span,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content),
            nodes,
            this,
            attributes: Attributes);
    }
}
