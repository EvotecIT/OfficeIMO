namespace OfficeIMO.Markdown;

/// <summary>
/// Represents an HTML comment preserved as a top-level Markdown block.
/// </summary>
public sealed class HtmlCommentBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Gets the raw HTML comment text, including the comment delimiters.</summary>
    public string Comment { get; }
    /// <summary>Source span for the opening <c>&lt;!--</c> marker when available.</summary>
    public MarkdownSourceSpan? OpeningMarkerSourceSpan { get; internal set; }
    /// <summary>Source span for the comment body when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; internal set; }
    /// <summary>Source span for the closing <c>--&gt;</c> marker when available.</summary>
    public MarkdownSourceSpan? ClosingMarkerSourceSpan { get; internal set; }

    /// <summary>Initializes a new instance of the <see cref="HtmlCommentBlock"/> class.</summary>
    /// <param name="comment">HTML comment content to preserve verbatim.</param>
    public HtmlCommentBlock(string comment) {
        Comment = comment ?? string.Empty;
    }

    string IMarkdownBlock.RenderMarkdown() => Comment;

    string IMarkdownBlock.RenderHtml() {
        var o = HtmlRenderContext.Options;
        var handling = o?.RawHtmlHandling ?? RawHtmlHandling.Allow;
        return handling switch {
            RawHtmlHandling.Allow => Comment,
            RawHtmlHandling.Escape => "<pre class=\"md-raw-html\"><code>" + System.Net.WebUtility.HtmlEncode(Comment) + "</code></pre>",
            RawHtmlHandling.Sanitize => RawHtmlSanitizer.Sanitize(Comment),
            _ => string.Empty
        };
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var commentParts = GetCommentParts();
        var children = new List<MarkdownSyntaxNode>();

        if (!OpeningMarkerSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(OpeningMarkerSourceSpan.Value))) {
            OpeningMarkerSourceSpan = GetSourceSpan(span, commentParts.OpeningStartIndex, commentParts.OpeningEndIndex);
        }

        if (!BodySourceSpan.HasValue || (span.HasValue && !span.Value.Contains(BodySourceSpan.Value))) {
            BodySourceSpan = commentParts.HasBody
                ? GetSourceSpan(span, commentParts.BodyStartIndex, commentParts.BodyEndIndex)
                : null;
        }

        if (!ClosingMarkerSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(ClosingMarkerSourceSpan.Value))) {
            ClosingMarkerSourceSpan = GetSourceSpan(span, commentParts.ClosingStartIndex, commentParts.ClosingEndIndex);
        }

        if (OpeningMarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlCommentOpeningMarker,
                OpeningMarkerSourceSpan,
                "<!--"));
        }

        if (BodySourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlCommentBody,
                BodySourceSpan,
                GetCommentBodyText(commentParts)));
        }

        if (ClosingMarkerSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlCommentClosingMarker,
                ClosingMarkerSourceSpan,
                "-->"));
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlComment, span, Comment, children, this);
    }

    private CommentParts GetCommentParts() {
        var openingStart = Comment.IndexOf("<!--", StringComparison.Ordinal);
        var closingStart = Comment.LastIndexOf("-->", StringComparison.Ordinal);
        if (openingStart < 0 || closingStart < openingStart + 4) {
            return default;
        }

        var bodyStart = openingStart + 4;
        var bodyEnd = closingStart - 1;

        while (bodyStart <= bodyEnd && Comment[bodyStart] == '\n') {
            bodyStart++;
        }

        while (bodyEnd >= bodyStart && Comment[bodyEnd] == '\n') {
            bodyEnd--;
        }

        return new CommentParts(
            openingStart,
            openingStart + 3,
            bodyStart,
            bodyEnd,
            closingStart,
            closingStart + 2);
    }

    private string GetCommentBodyText(CommentParts parts) {
        if (!parts.HasBody) {
            return string.Empty;
        }

        return Comment.Substring(parts.BodyStartIndex, parts.BodyEndIndex - parts.BodyStartIndex + 1);
    }

    private MarkdownSourceSpan? GetSourceSpan(MarkdownSourceSpan? blockSpan, int startIndex, int endIndex) {
        if (!blockSpan.HasValue || !blockSpan.Value.StartColumn.HasValue || startIndex < 0 || endIndex < startIndex || endIndex >= Comment.Length) {
            return null;
        }

        var start = GetPoint(blockSpan.Value, startIndex);
        var end = GetPoint(blockSpan.Value, endIndex);
        return new MarkdownSourceSpan(start.Line, start.Column, end.Line, end.Column);
    }

    private (int Line, int Column) GetPoint(MarkdownSourceSpan blockSpan, int index) {
        var line = blockSpan.StartLine;
        var column = blockSpan.StartColumn ?? 1;

        for (var i = 0; i < index && i < Comment.Length; i++) {
            if (Comment[i] == '\n') {
                line++;
                column = 1;
            } else {
                column++;
            }
        }

        return (line, column);
    }

    private readonly struct CommentParts {
        internal CommentParts(int openingStartIndex, int openingEndIndex, int bodyStartIndex, int bodyEndIndex, int closingStartIndex, int closingEndIndex) {
            OpeningStartIndex = openingStartIndex;
            OpeningEndIndex = openingEndIndex;
            BodyStartIndex = bodyStartIndex;
            BodyEndIndex = bodyEndIndex;
            ClosingStartIndex = closingStartIndex;
            ClosingEndIndex = closingEndIndex;
        }

        internal int OpeningStartIndex { get; }
        internal int OpeningEndIndex { get; }
        internal int BodyStartIndex { get; }
        internal int BodyEndIndex { get; }
        internal int ClosingStartIndex { get; }
        internal int ClosingEndIndex { get; }
        internal bool HasBody => BodyStartIndex >= 0 && BodyEndIndex >= BodyStartIndex;
    }
}
