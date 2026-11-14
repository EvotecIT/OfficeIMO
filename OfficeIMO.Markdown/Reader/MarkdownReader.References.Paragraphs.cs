namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    /// <summary>
    /// Whether a paragraph is open in front of a line during the reference pre-scan. The line that
    /// decides it is checked only when a definition follows it.
    /// </summary>
    private struct LazyParagraphContinuation {
        private readonly string[] _lines;
        private readonly MarkdownLoadOptions _options;
        private int _decidingLine;
        private bool _forced;

        internal LazyParagraphContinuation(string[] lines, MarkdownLoadOptions options) {
            _lines = lines;
            _options = options;
            _decidingLine = -1;
            _forced = false;
        }

        internal void Clear() {
            _decidingLine = -1;
            _forced = false;
        }

        internal void Set() {
            _decidingLine = -1;
            _forced = true;
        }

        internal void DecideAt(int lineIndex) {
            _decidingLine = lineIndex;
            _forced = false;
        }

        internal bool IsOpen => _forced
            || (_decidingLine >= 0 && IsReferenceDefinitionParagraphContinuationLine(_lines, _decidingLine, _options));
    }

    private static bool IsReferenceDefinitionParagraphContinuationLine(string[] lines, int index, MarkdownLoadOptions options) {
        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        if (IsAtxHeading(line, out _, out _) ||
            IsCodeFenceOpen(line, out _, out _, out _) ||
            StartsTable(lines, index, options) ||
            IsParagraphInterruptingThematicBreakLine(line) ||
            IsParagraphInterruptingUnorderedListLine(line) ||
            IsParagraphInterruptingOrderedListLine(line, options) ||
            (options.Callouts && IsCalloutHeader(line, options, out _, out _)) ||
            IsQuoteStarter(line) ||
            HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(line, options) ||
            IsCustomContainerOpeningLine(line, options) ||
            TryParseReferenceLinkDefinition(lines, index, options, out _, out _, out _, out _) ||
            IsFootnoteDefinitionStarterForParagraphInterruption(line, options) ||
            (options.StandaloneImageBlocks && IsImageLine(line))) {
            return false;
        }

        return true;
    }

    private static bool IsFootnoteDefinitionStarterForParagraphInterruption(string line, MarkdownLoadOptions options) {
        if (options?.Footnotes != true || string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') {
            leading++;
        }

        if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (!(trimmed.Length > 4 && trimmed[0] == '[' && trimmed[1] == '^')) {
            return false;
        }

        int rb = trimmed.IndexOf(']');
        return rb >= 2
               && rb + 1 < trimmed.Length
               && trimmed[rb + 1] == ':';
    }
}
