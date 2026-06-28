namespace OfficeIMO.Markdown;

/// <summary>
/// Raw HTML block passthrough.
/// </summary>
public sealed class HtmlRawBlock : MarkdownBlock, IMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Raw HTML content to emit.</summary>
    public string Html { get; }
    /// <summary>Source span for a recognized opening raw HTML tag when available.</summary>
    public MarkdownSourceSpan? OpeningTagSourceSpan { get; internal set; }
    /// <summary>Source span for the raw HTML body between recognized matching tags when available.</summary>
    public MarkdownSourceSpan? BodySourceSpan { get; internal set; }
    /// <summary>Source span for a recognized matching closing raw HTML tag when available.</summary>
    public MarkdownSourceSpan? ClosingTagSourceSpan { get; internal set; }

    /// <summary>Create a new raw HTML block.</summary>
    /// <param name="html">HTML fragment.</param>
    public HtmlRawBlock(string html) { Html = html ?? string.Empty; }
    string IMarkdownBlock.RenderMarkdown() => Html;
    string IMarkdownBlock.RenderHtml() {
        var o = HtmlRenderContext.Options;
        var handling = o?.RawHtmlHandling ?? RawHtmlHandling.Allow;
        return handling switch {
            RawHtmlHandling.Allow => o?.GitHubHtmlTagFilter == true ? GitHubHtmlTagFilter.Apply(Html) : Html,
            RawHtmlHandling.Escape => "<pre class=\"md-raw-html\"><code>" + System.Net.WebUtility.HtmlEncode(Html) + "</code></pre>",
            RawHtmlHandling.Sanitize => RawHtmlSanitizer.Sanitize(Html),
            _ => string.Empty
        };
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        if (!TryGetTagFrame(out var tagFrame)) {
            return new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlRaw, span, Html, associatedObject: this);
        }

        var children = new List<MarkdownSyntaxNode>();

        if (!OpeningTagSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(OpeningTagSourceSpan.Value))) {
            OpeningTagSourceSpan = HtmlBlockSourceSpanHelpers.GetSourceSpan(Html, span, tagFrame.OpeningStartIndex, tagFrame.OpeningEndIndex);
        }

        if (!BodySourceSpan.HasValue || (span.HasValue && !span.Value.Contains(BodySourceSpan.Value))) {
            BodySourceSpan = tagFrame.HasBody
                ? HtmlBlockSourceSpanHelpers.GetSourceSpan(Html, span, tagFrame.BodyStartIndex, tagFrame.BodyEndIndex)
                : null;
        }

        if (!ClosingTagSourceSpan.HasValue || (span.HasValue && !span.Value.Contains(ClosingTagSourceSpan.Value))) {
            ClosingTagSourceSpan = tagFrame.HasClosingTag
                ? HtmlBlockSourceSpanHelpers.GetSourceSpan(Html, span, tagFrame.ClosingStartIndex, tagFrame.ClosingEndIndex)
                : null;
        }

        if (OpeningTagSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlRawOpeningTag,
                OpeningTagSourceSpan,
                GetSlice(tagFrame.OpeningStartIndex, tagFrame.OpeningEndIndex)));
        }

        if (BodySourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlRawBody,
                BodySourceSpan,
                GetSlice(tagFrame.BodyStartIndex, tagFrame.BodyEndIndex)));
        }

        if (ClosingTagSourceSpan.HasValue) {
            children.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.HtmlRawClosingTag,
                ClosingTagSourceSpan,
                GetSlice(tagFrame.ClosingStartIndex, tagFrame.ClosingEndIndex)));
        }

        return new MarkdownSyntaxNode(MarkdownSyntaxKind.HtmlRaw, span, Html, children, this);
    }

    private bool TryGetTagFrame(out TagFrame tagFrame) {
        tagFrame = default;

        var openingStart = Html.IndexOf('<');
        if (openingStart < 0 || !TryReadOpeningTag(openingStart, out var tagName, out var openingEnd, out var isSelfClosing)) {
            return false;
        }

        if (isSelfClosing || !TryFindClosingTag(tagName, openingEnd, out var closingStart, out var closingEnd)) {
            tagFrame = new TagFrame(openingStart, openingEnd, -1, -1, -1, -1);
            return true;
        }

        var bodyStart = openingEnd + 1;
        var bodyEnd = closingStart - 1;

        while (bodyStart <= bodyEnd && Html[bodyStart] == '\n') {
            bodyStart++;
        }

        while (bodyEnd >= bodyStart && Html[bodyEnd] == '\n') {
            bodyEnd--;
        }

        tagFrame = new TagFrame(openingStart, openingEnd, bodyStart, bodyEnd, closingStart, closingEnd);
        return true;
    }

    private bool TryReadOpeningTag(int openingStart, out string tagName, out int openingEnd, out bool isSelfClosing) {
        tagName = string.Empty;
        openingEnd = -1;
        isSelfClosing = false;

        var tagNameStart = openingStart + 1;
        if (tagNameStart >= Html.Length || !IsAsciiLetter(Html[tagNameStart])) {
            return false;
        }

        var tagNameEnd = tagNameStart + 1;
        while (tagNameEnd < Html.Length && IsTagNameCharacter(Html[tagNameEnd])) {
            tagNameEnd++;
        }

        tagName = Html.Substring(tagNameStart, tagNameEnd - tagNameStart);
        openingEnd = FindTagEnd(openingStart);
        if (openingEnd < 0) {
            return false;
        }

        var beforeEnd = openingEnd - 1;
        while (beforeEnd > openingStart && char.IsWhiteSpace(Html[beforeEnd])) {
            beforeEnd--;
        }

        isSelfClosing = beforeEnd > openingStart && Html[beforeEnd] == '/';
        return true;
    }

    private bool TryFindClosingTag(string tagName, int openingEnd, out int closingStart, out int closingEnd) {
        closingStart = -1;
        closingEnd = -1;

        var search = "</" + tagName;
        var startIndex = Html.Length - 1;
        while (startIndex > openingEnd) {
            var candidate = Html.LastIndexOf(search, startIndex, StringComparison.OrdinalIgnoreCase);
            if (candidate <= openingEnd) {
                return false;
            }

            var afterName = candidate + search.Length;
            if (IsTagNameBoundary(afterName)) {
                var candidateEnd = FindTagEnd(candidate);
                if (candidateEnd >= 0) {
                    closingStart = candidate;
                    closingEnd = candidateEnd;
                    return true;
                }
            }

            startIndex = candidate - 1;
        }

        return false;
    }

    private int FindTagEnd(int tagStart) {
        var quote = '\0';
        for (var i = tagStart + 1; i < Html.Length; i++) {
            var ch = Html[i];
            if (quote != '\0') {
                if (ch == quote) {
                    quote = '\0';
                }

                continue;
            }

            if (ch == '"' || ch == '\'') {
                quote = ch;
            } else if (ch == '>') {
                return i;
            }
        }

        return -1;
    }

    private string GetSlice(int startIndex, int endIndex) =>
        startIndex >= 0 && endIndex >= startIndex && endIndex < Html.Length
            ? Html.Substring(startIndex, endIndex - startIndex + 1)
            : string.Empty;

    private bool IsTagNameBoundary(int index) =>
        index >= Html.Length || Html[index] == '>' || char.IsWhiteSpace(Html[index]);

    private static bool IsTagNameCharacter(char ch) =>
        IsAsciiLetter(ch) || char.IsDigit(ch) || ch == '-' || ch == ':';

    private static bool IsAsciiLetter(char ch) =>
        ch is >= 'A' and <= 'Z' or >= 'a' and <= 'z';

    private readonly struct TagFrame {
        internal TagFrame(int openingStartIndex, int openingEndIndex, int bodyStartIndex, int bodyEndIndex, int closingStartIndex, int closingEndIndex) {
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
        internal bool HasClosingTag => ClosingStartIndex >= 0 && ClosingEndIndex >= ClosingStartIndex;
    }
}
