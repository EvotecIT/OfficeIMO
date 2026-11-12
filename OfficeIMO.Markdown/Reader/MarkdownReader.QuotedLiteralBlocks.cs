namespace OfficeIMO.Markdown;

/// <summary>The quoted content (one '>' level stripped) of one run of quote lines in a source line array.</summary>
internal sealed class QuotedContentRun {
    private readonly string[] _source;
    private readonly int _continuationIndent;
    private readonly int _start;

    internal QuotedContentRun(string[] source, int continuationIndent, int start, string[] content) {
        _source = source;
        _continuationIndent = continuationIndent;
        _start = start;
        Content = content;
    }

    internal string[] Content { get; }

    /// <summary>Returns true when line <paramref name="index"/> of <paramref name="source"/> lies in this run.</summary>
    internal bool Covers(string[] source, int continuationIndent, int index) =>
        ReferenceEquals(_source, source)
        && _continuationIndent == continuationIndent
        && index >= _start
        && index - _start < Content.Length;

    internal int OffsetOf(int index) => index - _start;
}

internal static partial class MarkdownReader {
    // Marks top-level quotes, which have no list-item continuation indent.
    private const int TopLevelQuote = -1;

    /// <summary>
    /// Tracks fenced code and HTML blocks inside quoted content, whose lines are literal text: a callout-like line in
    /// them does not start a new callout, and an unquoted line does not continue them lazily.
    /// </summary>
    private sealed class QuotedLiteralBlockState {
        private readonly MarkdownReaderState _state;
        private bool _inFence;
        private char _fenceChar;
        private int _fenceLength;
        private int _htmlLinesLeft;
        private bool _htmlReachesRunEnd;

        internal QuotedLiteralBlockState(MarkdownReaderState state) {
            _state = state;
        }

        /// <summary>
        /// True while a quoted fence or HTML block is open. An HTML block that runs to the last quoted line stays open
        /// there: whether or not it closed, an unquoted line after it cannot continue it lazily, since only paragraphs do.
        /// </summary>
        internal bool IsOpen => _inFence || _htmlLinesLeft > 0 || _htmlReachesRunEnd;

        /// <summary>Returns true when the line belongs to an open fence or HTML block, closing the fence on its closing line.</summary>
        internal bool Continue(string content) {
            if (_htmlLinesLeft > 0) {
                _htmlLinesLeft--;
                return true;
            }

            if (!_inFence) return false;
            if (IsCodeFenceClose(content, _fenceChar, _fenceLength)) _inFence = false;
            return true;
        }

        /// <summary>Opens a fence when fenced code is enabled and the line starts one.</summary>
        internal bool TryOpenFence(string content, MarkdownLoadOptions options) {
            if (!options.FencedCode || !IsCodeFenceOpen(content, out _, out _fenceChar, out _fenceLength)) return false;
            _inFence = true;
            return true;
        }

        /// <summary>Opens an HTML block that starts on the top-level quote line at <paramref name="index"/>.</summary>
        internal void TryOpenTopLevelHtmlBlock(string content, string[] lines, int index, MarkdownLoadOptions options) {
            if (!StartsHtmlBlock(content, options)) return;
            var run = _state.QuotedContentRun;
            if (run == null || !run.Covers(lines, TopLevelQuote, index)) {
                run = new QuotedContentRun(lines, TopLevelQuote, index, CollectTopLevelQuotedContent(lines, index));
                _state.QuotedContentRun = run;
            }

            OpenHtmlBlock(run, index, options);
        }

        /// <summary>Opens an HTML block that starts on the list-item quote line at <paramref name="index"/>.</summary>
        internal void TryOpenNestedHtmlBlock(string content, string[] lines, int index, int continuationIndent, MarkdownLoadOptions options) {
            if (!StartsHtmlBlock(content, options)) return;
            var run = _state.QuotedContentRun;
            if (run == null || !run.Covers(lines, continuationIndent, index)) {
                run = new QuotedContentRun(lines, continuationIndent, index, CollectNestedQuotedContent(lines, index, continuationIndent));
                _state.QuotedContentRun = run;
            }

            OpenHtmlBlock(run, index, options);
        }

        // HTML blocks only matter for callout boundaries, so they are not measured when callouts are off.
        private static bool StartsHtmlBlock(string content, MarkdownLoadOptions options) =>
            options.Callouts && HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(content, options);

        private void OpenHtmlBlock(QuotedContentRun run, int index, MarkdownLoadOptions options) {
            int offset = run.OffsetOf(index);
            if (HtmlBlockParser.TryGetHtmlBlockLineCount(run.Content, offset, options, out int lineCount)) {
                _htmlLinesLeft = lineCount - 1;
                _htmlReachesRunEnd = offset + lineCount >= run.Content.Length;
            }
        }
    }

    // The quoted content (one '>' level stripped) of the consecutive top-level quote lines starting at start.
    private static string[] CollectTopLevelQuotedContent(string[] lines, int start) {
        var content = new System.Collections.Generic.List<string>();
        for (int k = start; k < lines.Length; k++) {
            string line = lines[k] ?? string.Empty;
            if (CountLeadingIndentColumns(line) > 3 || !line.TrimStart().StartsWith(">", StringComparison.Ordinal)) break;
            content.Add(StripSingleQuoteMarker(ExpandWhitespaceAfterContainerMarker(line, line.IndexOf('>') + 1)));
        }

        return content.ToArray();
    }

    // The quoted content of the consecutive quote lines in a list item, from start, at the item's continuation indent.
    private static string[] CollectNestedQuotedContent(string[] lines, int start, int continuationIndent) {
        var content = new System.Collections.Generic.List<string>();
        for (int k = start; k < lines.Length; k++) {
            string raw = lines[k] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw) || CountLeadingIndentColumns(raw) < continuationIndent) break;
            string part = StripLeadingIndentColumns(raw, continuationIndent);
            if (CountLeadingIndentColumns(part) > 3 || !part.TrimStart().StartsWith(">", StringComparison.Ordinal)) break;
            content.Add(StripSingleQuoteMarker(part));
        }

        return content.ToArray();
    }
}
