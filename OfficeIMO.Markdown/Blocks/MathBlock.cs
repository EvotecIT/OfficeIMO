namespace OfficeIMO.Markdown;

/// <summary>
/// Display math block delimited by <c>$$</c> when <see cref="MarkdownLoadOptions.Math"/> is enabled.
/// </summary>
/// <remarks>
/// The block carries TeX source only. Converters lay it out through their own math engines, and HTML output keeps
/// <c>\[...\]</c> delimiters for client-side renderers such as KaTeX or MathJax. Display math from
/// <c>```math</c> fences and from LaTeX or AsciiDoc projections uses <see cref="SemanticFencedBlock"/> with the
/// <c>math</c> semantic kind.
/// </remarks>
public sealed class MathBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Creates a display math block.</summary>
    /// <param name="content">TeX source without the <c>$$</c> delimiters.</param>
    public MathBlock(string content) {
        Content = (content ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
    }

    /// <summary>TeX source without the <c>$$</c> delimiters.</summary>
    public string Content { get; }

    /// <summary>Source span for the TeX source when parsed from Markdown.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; private set; }

    /// <summary>Source span for the opening <c>$$</c> delimiter when parsed from Markdown.</summary>
    public MarkdownSourceSpan? OpeningDelimiterSourceSpan { get; private set; }

    /// <summary>Source span for the closing <c>$$</c> delimiter when parsed from Markdown.</summary>
    public MarkdownSourceSpan? ClosingDelimiterSourceSpan { get; private set; }

    /// <summary>Whether both delimiters and the source share one line (<c>$$x$$</c>).</summary>
    internal bool IsSingleLine { get; set; }

    internal int IndentColumns { get; set; }

    string IMarkdownBlock.RenderMarkdown() =>
        IsSingleLine && Content.IndexOf('\n') < 0
            ? "$$" + Content + "$$"
            : "$$\n" + Content + "\n$$";

    string IMarkdownBlock.RenderHtml() =>
        "<div class=\"math math-display\">\\[" + HtmlTextEncoder.Encode(Content, HtmlRenderContext.Options) + "\\]</div>";

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        ComputeSourceSpans(span);
        var children = new List<MarkdownSyntaxNode>(3);
        if (OpeningDelimiterSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.MathBlockDelimiter, OpeningDelimiterSourceSpan, "$$"));
        }

        children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.MathBlockContent, ContentSourceSpan, Content));
        if (ClosingDelimiterSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(MarkdownSyntaxKind.MathBlockDelimiter, ClosingDelimiterSourceSpan, "$$"));
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.MathBlock, span, Content, children, this);
    }

    private void ComputeSourceSpans(MarkdownSourceSpan? span) {
        OpeningDelimiterSourceSpan = null;
        ClosingDelimiterSourceSpan = null;
        ContentSourceSpan = null;
        if (!span.HasValue) {
            return;
        }

        MarkdownSourceSpan value = span.Value;
        if (!value.StartColumn.HasValue) {
            ContentSourceSpan = value;
            return;
        }

        int startColumn = value.StartColumn.Value + IndentColumns;
        int? startOffset = value.StartOffset.HasValue ? value.StartOffset.Value + IndentColumns : null;
        OpeningDelimiterSourceSpan = new MarkdownSourceSpan(value.StartLine, startColumn, value.StartLine, startColumn + 1,
            startOffset, startOffset + 1);
        if (IsSingleLine) {
            int contentStart = startColumn + 2;
            if (Content.Length > 0) {
                ContentSourceSpan = new MarkdownSourceSpan(value.StartLine, contentStart, value.StartLine, contentStart + Content.Length - 1,
                    startOffset + 2, startOffset + 1 + Content.Length);
            }

            int closingStart = contentStart + Content.Length;
            ClosingDelimiterSourceSpan = new MarkdownSourceSpan(value.StartLine, closingStart, value.StartLine, closingStart + 1,
                startOffset + 2 + Content.Length, startOffset + 3 + Content.Length);
            return;
        }

        int contentLines = Content.Length == 0 ? 0 : Content.Split('\n').Length;
        if (contentLines > 0) {
            int lastLength = Content.Length - (Content.LastIndexOf('\n') + 1);
            ContentSourceSpan = new MarkdownSourceSpan(value.StartLine + 1, 1, value.StartLine + contentLines, Math.Max(1, lastLength));
        }

        int closingLine = value.StartLine + contentLines + 1;
        if (closingLine <= value.EndLine) {
            ClosingDelimiterSourceSpan = new MarkdownSourceSpan(closingLine, startColumn, closingLine, startColumn + 1);
        }
    }
}
