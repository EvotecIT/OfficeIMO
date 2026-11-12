namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Tracks fenced code and HTML blocks inside quoted content, whose lines are literal text: a callout-like line in
    /// them does not start a new callout, and an unquoted line does not continue them lazily.
    /// </summary>
    private sealed class QuotedLiteralBlockState {
        private bool _inFence;
        private char _fenceChar;
        private int _fenceLength;
        private int _htmlLinesLeft;

        // The quoted content of one run of quote lines, collected once and reused for every HTML block in the run.
        private string[]? _run;
        private int _runStart;

        /// <summary>True while a quoted fence or HTML block is open.</summary>
        internal bool IsOpen => _inFence || _htmlLinesLeft > 0;

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
            if (!RunCovers(index)) StartRun(CollectTopLevelQuotedContent(lines, index), index);
            OpenHtmlBlock(index, options);
        }

        /// <summary>Opens an HTML block that starts on the list-item quote line at <paramref name="index"/>.</summary>
        internal void TryOpenNestedHtmlBlock(string content, string[] lines, int index, int continuationIndent, MarkdownLoadOptions options) {
            if (!StartsHtmlBlock(content, options)) return;
            if (!RunCovers(index)) StartRun(CollectNestedQuotedContent(lines, index, continuationIndent), index);
            OpenHtmlBlock(index, options);
        }

        // HTML blocks only matter for callout boundaries, so they are not measured when callouts are off.
        private static bool StartsHtmlBlock(string content, MarkdownLoadOptions options) =>
            options.Callouts && HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(content, options);

        private bool RunCovers(int index) => _run != null && index >= _runStart && index - _runStart < _run.Length;

        private void StartRun(string[] run, int start) {
            _run = run;
            _runStart = start;
        }

        private void OpenHtmlBlock(int index, MarkdownLoadOptions options) {
            if (HtmlBlockParser.TryGetHtmlBlockLineCount(_run!, index - _runStart, options, out int lineCount)) {
                _htmlLinesLeft = lineCount - 1;
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
