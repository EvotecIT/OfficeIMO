namespace OfficeIMO.Markdown;

/// <summary>
/// Fenced code block with optional caption. Fence length adapts to backticks inside the content.
/// </summary>
public sealed class CodeBlock : MarkdownBlock, IMarkdownBlock, ICaptionable, ISyntaxMarkdownBlock {
    /// <summary>Parsed primary fence language token (for example <c>csharp</c> or <c>bash</c>).</summary>
    public string Language { get; }
    /// <summary>Full fenced-code info string as it appeared after the opening fence marker.</summary>
    public string InfoString { get; }
    /// <summary>Structured fenced-code info metadata.</summary>
    public MarkdownCodeFenceInfo FenceInfo { get; }
    /// <summary>Code contents.</summary>
    public string Content { get; }
    /// <summary>Optional caption shown under the block.</summary>
    public string? Caption { get; set; }
    internal bool IsFenced { get; }
    internal int FenceIndentColumns { get; private set; }
    internal char FenceChar { get; private set; } = '`';
    internal int FenceLength { get; private set; } = 3;
    internal int FenceInfoPaddingColumns { get; private set; }
    internal bool HasClosingFence { get; private set; } = true;
    internal int ClosingFenceIndentColumns { get; private set; }
    internal int ClosingFenceLength { get; private set; } = 3;
    /// <summary>Source span for the opening fence token when parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; internal set; }
    /// <summary>Source span for the closing fence token when parsed from a closed fenced source block.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; internal set; }

    /// <summary>Create a code block with an optional fenced-code info string.</summary>
    public CodeBlock(string language, string content) : this(language, content, isFenced: true) {
    }

    internal CodeBlock(string language, string content, bool isFenced) {
        FenceInfo = MarkdownCodeFenceInfo.Parse(language);
        InfoString = FenceInfo.InfoString;
        Language = FenceInfo.Language;
        Content = NormalizeLineEndings(content);
        IsFenced = isFenced;
        SetAttributes(MarkdownAttributeSet.Create(FenceInfo.ElementId, FenceInfo.Classes, FenceInfo.Attributes));
    }

    internal void SetFenceSourceInfo(
        int fenceIndentColumns,
        int fenceLength,
        int infoPaddingColumns,
        char fenceChar = '`',
        bool hasClosingFence = true,
        int closingFenceIndentColumns = 0,
        int closingFenceLength = 3) {
        FenceIndentColumns = Math.Max(0, fenceIndentColumns);
        FenceChar = fenceChar == '~' ? '~' : '`';
        FenceLength = Math.Max(3, fenceLength);
        FenceInfoPaddingColumns = Math.Max(0, infoPaddingColumns);
        HasClosingFence = hasClosingFence;
        ClosingFenceIndentColumns = Math.Max(0, closingFenceIndentColumns);
        ClosingFenceLength = Math.Max(3, closingFenceLength);
    }

    internal void SetFenceTokenSourceSpans(MarkdownSourceSpan? openingFenceSourceSpan, MarkdownSourceSpan? closingFenceSourceSpan) {
        OpeningFenceSourceSpan = openingFenceSourceSpan;
        ClosingFenceSourceSpan = closingFenceSourceSpan;
    }

    private static string NormalizeLineEndings(string? content) {
        if (string.IsNullOrEmpty(content)) {
            return string.Empty;
        }

        return content!
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        string fence = MarkdownFence.BuildSafeFence(Content);

        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{fence}{InfoString}");
        sb.AppendLine(Content);
        sb.AppendLine(fence);
        if (!string.IsNullOrWhiteSpace(Caption)) sb.AppendLine("_" + Caption + "_");
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var options = HtmlRenderContext.Options;
        var overridden = options?.CodeBlockHtmlRenderer?.Invoke(this, options);
        if (overridden != null) {
            return overridden;
        }

        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{HtmlTextEncoder.Encode(Language, options)}\"";
        string code = HtmlTextEncoder.Encode(Content, options);
        if (code.Length > 0) {
            // CommonMark-style HTML keeps the terminating line break inside <code>
            // for multi-line block code, even though the stored model content does not.
            code += "\n";
        }
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{HtmlTextEncoder.Encode(Caption!, options)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode>();
        OpeningFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetOpeningFenceSpan(span, IsFenced, FenceIndentColumns, FenceLength);
        ClosingFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetClosingFenceSpan(span, IsFenced, Content, HasClosingFence, ClosingFenceIndentColumns, ClosingFenceLength);

        if (OpeningFenceSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceOpening,
                OpeningFenceSourceSpan,
                new string(FenceChar, FenceLength)));
        }

        var infoSpan = MarkdownFencedBlockSourceSpans.GetInfoSpan(span, IsFenced, InfoString, FenceIndentColumns, FenceLength, FenceInfoPaddingColumns);
        if (infoSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                infoSpan,
                InfoString));
        }

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeContent,
            MarkdownFencedBlockSourceSpans.GetContentSpan(span, IsFenced, Content),
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content)));

        if (ClosingFenceSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceClosing,
                ClosingFenceSourceSpan,
                new string(FenceChar, ClosingFenceLength)));
        }

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeBlock,
            span,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content),
            nodes,
            this,
            attributes: Attributes);
    }
}
