namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Tracks fenced code and HTML blocks inside quoted content, whose lines are literal text: a callout-like line in
    /// them does not start a new callout.
    /// </summary>
    private sealed class QuotedLiteralBlockState {
        private bool _inFence;
        private char _fenceChar;
        private int _fenceLength;
        private int _htmlLinesLeft;

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

        /// <summary>Returns true when HTML blocks are enabled and the line starts one that can interrupt a paragraph.</summary>
        internal static bool StartsHtmlBlock(string content, MarkdownLoadOptions options) =>
            HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(content, options);

        /// <summary>Marks the rest of the HTML block that starts at the first of <paramref name="quotedContent"/> as literal.</summary>
        internal void OpenHtmlBlock(string[] quotedContent, MarkdownLoadOptions options) {
            if (HtmlBlockParser.TryGetHtmlBlockLineCount(quotedContent, 0, options, out int lineCount)) {
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
