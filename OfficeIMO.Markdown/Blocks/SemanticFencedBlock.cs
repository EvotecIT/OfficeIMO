namespace OfficeIMO.Markdown;

/// <summary>
/// First-class AST node for fenced blocks whose language maps to host-defined semantics
/// such as diagrams, charts, data views, or other non-code contracts.
/// </summary>
public sealed class SemanticFencedBlock : IMarkdownBlock, ICaptionable, ISyntaxMarkdownBlock {
    /// <summary>Create a semantic fenced block.</summary>
    public SemanticFencedBlock(string semanticKind, string language, string content, string? caption = null)
        : this(semanticKind, language, content, caption, isFenced: true) {
    }

    internal SemanticFencedBlock(string semanticKind, string language, string content, string? caption, bool isFenced) {
        SemanticKind = string.IsNullOrWhiteSpace(semanticKind) ? MarkdownSemanticKinds.Custom : semanticKind.Trim();
        Language = language ?? string.Empty;
        Content = content ?? string.Empty;
        Caption = caption;
        IsFenced = isFenced;
    }

    /// <summary>Host-defined semantic contract for this block (for example <c>chart</c> or <c>mermaid</c>).</summary>
    public string SemanticKind { get; }

    /// <summary>Original fence language / info string.</summary>
    public string Language { get; }

    /// <summary>Raw fenced payload.</summary>
    public string Content { get; }

    /// <summary>Optional caption shown under the block.</summary>
    public string? Caption { get; set; }

    internal bool IsFenced { get; }

    string IMarkdownBlock.RenderMarkdown() {
        string fence = MarkdownFence.BuildSafeFence(Content);

        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{fence}{Language}");
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

        var codeFallback = options?.CodeBlockHtmlRenderer?.Invoke(new CodeBlock(Language, Content) {
            Caption = Caption
        }, options);
        if (codeFallback != null) {
            return codeFallback;
        }

        string lang = string.IsNullOrEmpty(Language) ? string.Empty : $" class=\"language-{System.Net.WebUtility.HtmlEncode(Language)}\"";
        string code = System.Net.WebUtility.HtmlEncode(Content);
        if (code.Length > 0) {
            code += "\n";
        }

        string caption = string.IsNullOrWhiteSpace(Caption) ? string.Empty : $"<div class=\"caption\">{System.Net.WebUtility.HtmlEncode(Caption!)}</div>";
        return $"<pre><code{lang}>{code}</code></pre>{caption}";
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var nodes = new List<MarkdownSyntaxNode> {
            new MarkdownSyntaxNode(MarkdownSyntaxKind.FenceSemanticKind, literal: SemanticKind)
        };

        if (span.HasValue && IsFenced && !string.IsNullOrEmpty(Language)) {
            nodes.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.CodeFenceInfo,
                new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartLine),
                Language));
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
            MarkdownSyntaxKind.SemanticFencedBlock,
            span,
            MarkdownBlockSyntaxBuilder.NormalizeSyntaxLiteralLineEndings(Content),
            nodes);
    }
}
