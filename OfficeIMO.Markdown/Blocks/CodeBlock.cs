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
    internal int FenceInfoPaddingCharacters { get; private set; }
    internal bool HasClosingFence { get; private set; } = true;
    internal int ClosingFenceIndentColumns { get; private set; }
    internal int ClosingFenceLength { get; private set; } = 3;
    internal int? FencedContentLineCount { get; private set; }
    /// <summary>Source span for the opening fence token when parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; internal set; }
    /// <summary>Source span for the fenced-code info string when parsed from a fenced source block.</summary>
    public MarkdownSourceSpan? InfoStringSourceSpan { get; internal set; }
    /// <summary>Source span for the code payload when source-backed.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; internal set; }
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
        if (FenceInfo.HasExplicitAttributes) {
            SetAttributes(MarkdownAttributeSet.Create(FenceInfo.ElementId, FenceInfo.Classes, FenceInfo.Attributes));
        }
    }

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
        var standaloneAttributes = GetStandaloneRenderedAttributes();
        if (!standaloneAttributes.IsEmpty) {
            sb.AppendLine(MarkdownAttributeBlockRenderer.RenderInlineTrailing(standaloneAttributes));
        }

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

        bool renderGenericAttributesOnCode = ShouldRenderGenericAttributesOnCode();
        string attrs = renderGenericAttributesOnCode ? string.Empty : MarkdownHtmlAttributes.Render(Attributes, options);
        string lang = renderGenericAttributesOnCode
            ? RenderCodeAttributes(options)
            : string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{HtmlTextEncoder.Encode(Language, options)}\"";
        string code = HtmlTextEncoder.Encode(Content, options);
        if (code.Length > 0) {
            // CommonMark-style HTML keeps the terminating line break inside <code>
            // for multi-line block code, even though the stored model content does not.
            code += "\n";
        }
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{HtmlTextEncoder.Encode(Caption!, options)}</div>";
        return $"<pre{attrs}><code{lang}>{code}</code></pre>{caption}";
    }

    private bool ShouldRenderStandaloneGenericAttributes() =>
        !Attributes.IsEmpty && MarkdownGenericAttributeSourceSpans.GetSourceSpan(this).HasValue;

    private MarkdownAttributeSet GetStandaloneRenderedAttributes() {
        if (!ShouldRenderStandaloneGenericAttributes()) {
            return MarkdownAttributeSet.Empty;
        }

        if (!FenceInfo.HasExplicitAttributes) {
            return Attributes;
        }

        var explicitAttributes = FenceInfo.GenericAttributes;
        var elementId = string.Equals(Attributes.ElementId, explicitAttributes.ElementId, StringComparison.Ordinal)
            ? null
            : Attributes.ElementId;
        var classes = new List<string>();
        for (int i = 0; i < Attributes.Classes.Count; i++) {
            var className = Attributes.Classes[i];
            if (!explicitAttributes.HasClass(className)) {
                classes.Add(className);
            }
        }

        var attributes = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var attribute in Attributes.Attributes) {
            if (explicitAttributes.TryGetAttribute(attribute.Key, out var explicitValue)
                && string.Equals(attribute.Value, explicitValue, StringComparison.Ordinal)) {
                continue;
            }

            attributes[attribute.Key] = attribute.Value;
        }

        return MarkdownAttributeSet.Create(elementId, classes, attributes);
    }

    private bool ShouldRenderGenericAttributesOnCode() =>
        !Attributes.IsEmpty && (FenceInfo.HasExplicitAttributes || ShouldRenderStandaloneGenericAttributes());

    private string RenderCodeAttributes(HtmlOptions? options) {
        if (Attributes.IsEmpty && string.IsNullOrEmpty(Language)) {
            return string.Empty;
        }

        var classes = new List<string>();
        bool renderFenceInfoAttributesFirst = FenceInfo.HasExplicitAttributes && !ShouldRenderStandaloneGenericAttributes();
        if (!renderFenceInfoAttributesFirst && !string.IsNullOrEmpty(Language)) {
            classes.Add("language-" + Language);
        }

        var sourceAttributes = renderFenceInfoAttributesFirst ? GetFenceInfoCodeAttributes() : Attributes;
        for (int i = 0; i < sourceAttributes.Classes.Count; i++) {
            classes.Add(sourceAttributes.Classes[i]);
        }

        if (renderFenceInfoAttributesFirst && !string.IsNullOrEmpty(Language)) {
            classes.Add("language-" + Language);
        }

        var codeAttributes = MarkdownAttributeSet.Create(
            sourceAttributes.ElementId,
            classes,
            sourceAttributes.Attributes);

        return MarkdownHtmlAttributes.Render(codeAttributes, options);
    }

    private MarkdownAttributeSet GetFenceInfoCodeAttributes() {
        if (!FenceInfo.HasExplicitAttributes) {
            return Attributes;
        }

        return MarkdownAttributeSet.Create(
            FenceInfo.ElementId,
            FenceInfo.Classes,
            FenceInfo.GenericAttributes.Attributes);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode>();
        var fenceSpan = GetFenceSourceSpan(span);
        OpeningFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetOpeningFenceSpan(fenceSpan, IsFenced, FenceIndentColumns, FenceLength);
        ClosingFenceSourceSpan = MarkdownFencedBlockSourceSpans.GetClosingFenceSpan(fenceSpan, IsFenced, Content, HasClosingFence, ClosingFenceIndentColumns, ClosingFenceLength, FencedContentLineCount);

        if (OpeningFenceSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceOpening,
                OpeningFenceSourceSpan,
                new string(FenceChar, FenceLength)));
        }

        InfoStringSourceSpan = MarkdownFencedBlockSourceSpans.GetInfoSpan(fenceSpan, IsFenced, InfoString, FenceIndentColumns, FenceLength, FenceInfoPaddingColumns, FenceInfoPaddingCharacters);
        if (InfoStringSourceSpan.HasValue) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                InfoStringSourceSpan,
                InfoString));
        }

        ContentSourceSpan = MarkdownFencedBlockSourceSpans.GetContentSpan(fenceSpan, IsFenced, Content, FencedContentLineCount);
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
            MarkdownSyntaxKind.CodeBlock,
            span,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content),
            nodes,
            this,
            attributes: Attributes);
    }

    private MarkdownSourceSpan? GetFenceSourceSpan(MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return span;
        }

        var attributeSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(this);
        if (!attributeSpan.HasValue) {
            return span;
        }

        var value = span.Value;
        var fenceStartLine = attributeSpan.Value.EndLine + 1;
        if (fenceStartLine > value.EndLine) {
            return span;
        }

        if (!value.EndColumn.HasValue) {
            return new MarkdownSourceSpan(fenceStartLine, value.EndLine);
        }

        return new MarkdownSourceSpan(fenceStartLine, 1, value.EndLine, value.EndColumn.Value);
    }
}
