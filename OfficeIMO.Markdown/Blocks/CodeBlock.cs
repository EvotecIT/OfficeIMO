namespace OfficeIMO.Markdown;

/// <summary>
/// Fenced code block with optional caption. Fence length adapts to backticks inside the content.
/// </summary>
public sealed class CodeBlock : IMarkdownBlock, ICaptionable, ISyntaxMarkdownBlock {
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

    /// <summary>Create a code block with an optional fenced-code info string.</summary>
    public CodeBlock(string language, string content) : this(language, content, isFenced: true) {
    }

    internal CodeBlock(string language, string content, bool isFenced) {
        FenceInfo = MarkdownCodeFenceInfo.Parse(language);
        InfoString = FenceInfo.InfoString;
        Language = FenceInfo.Language;
        Content = NormalizeLineEndings(content);
        IsFenced = isFenced;
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

        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\"";
        string code = System.Net.WebUtility.HtmlEncode(Content);
        if (code.Length > 0) {
            // CommonMark-style HTML keeps the terminating line break inside <code>
            // for multi-line block code, even though the stored model content does not.
            code += "\n";
        }
        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode>();
        if (span.HasValue && IsFenced && !string.IsNullOrEmpty(InfoString)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartLine),
                InfoString));
        }

        MarkdownSourceSpan? contentSpan;
        if (span.HasValue) {
            if (IsFenced) {
                contentSpan = span.Value.EndLine > span.Value.StartLine + 1
                    ? new MarkdownSourceSpan(span.Value.StartLine + 1, span.Value.EndLine - 1)
                    : null;
            } else {
                contentSpan = span.Value;
            }
        } else {
            contentSpan = null;
        }

        nodes.Add(new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeContent,
            contentSpan,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content)));

        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.CodeBlock,
            span,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content),
            nodes);
    }
}
