namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HrParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            // HR is independent of options; it's safe and tiny, but follow Paragraphs toggle for symmetry
            if (!options.Paragraphs) return false;
            if (!LooksLikeHr(lines[i])) return false;
            if (TryTakePendingGenericAttributeBlock(state, out var pending)
                && options.Headings
                && TryGetSetextHeadingUnderlineLevel(lines[i], out var level)) {
                var heading = new HeadingBlock(level, new InlineSequence());
                heading.SetAttributes(pending.Attributes);
                MarkdownGenericAttributeSourceSpans.Set(heading, pending.SourceText, pending.SourceSpan);

                var underline = lines[i] ?? string.Empty;
                var trimmedUnderline = underline.Trim();
                var markerStartColumn = underline.IndexOf(trimmedUnderline, StringComparison.Ordinal) + 1;
                var markerEndColumn = markerStartColumn + trimmedUnderline.Length - 1;
                var absoluteMarkerLine = state.SourceLineOffset + i + 1;
                heading.SetLevelSourceInfo(
                    0,
                    markerStartColumn,
                    markerEndColumn);
                heading.SetSetextUnderlineMarkerSourceInfo(
                    0,
                    markerStartColumn,
                    markerEndColumn,
                    trimmedUnderline,
                    CreateSpan(state, absoluteMarkerLine, markerStartColumn, absoluteMarkerLine, markerEndColumn));
                doc.Add(heading);
                i++;
                return true;
            }

            var horizontalRule = new HorizontalRuleBlock();
            SetThematicBreakMarkerSource(horizontalRule, lines[i], state.SourceLineOffset + i + 1, state);
            doc.Add(horizontalRule);
            i++; return true;
        }
    }

    private static bool LooksLikeHr(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) > 3) return false;
        var t = line.Trim();
        // Allow spaces between characters, e.g., "- - -". Only '-', '*', '_' are valid.
        char kind = '\0'; int count = 0;
        for (int k = 0; k < t.Length; k++) {
            char ch = t[k];
            if (ch == ' ' || ch == '\t') continue;
            if (ch == '-' || ch == '*' || ch == '_') {
                if (kind == '\0') kind = ch; else if (kind != ch) return false;
                count++;
            } else return false;
        }
        return count >= 3;
    }

    private static bool IsParagraphInterruptingThematicBreakLine(string line) =>
        LooksLikeHr(line) && !LooksLikeSetextHeadingUnderline(line);

    private static void SetThematicBreakMarkerSource(
        HorizontalRuleBlock horizontalRule,
        string line,
        int absoluteLineNumber,
        MarkdownReaderState state) {
        if (horizontalRule == null || !TryGetTrimmedMarkerBounds(line, out var startIndex, out var endIndex, out var startColumn, out var endColumn)) {
            return;
        }

        horizontalRule.MarkerText = line.Substring(startIndex, endIndex - startIndex + 1);
        horizontalRule.MarkerSourceSpan = CreateSpan(state, absoluteLineNumber, startColumn, absoluteLineNumber, endColumn);
    }

    private static bool TryGetTrimmedMarkerBounds(
        string line,
        out int startIndex,
        out int endIndex,
        out int startColumn,
        out int endColumn) {
        startIndex = -1;
        endIndex = -1;
        startColumn = 1;
        endColumn = 1;
        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        var column = 1;
        for (var i = 0; i < line.Length; i++) {
            var ch = line[i];
            var currentColumn = column;
            if (ch != ' ' && ch != '\t') {
                if (startIndex < 0) {
                    startIndex = i;
                    startColumn = currentColumn;
                }

                endIndex = i;
                endColumn = currentColumn;
            }

            column = ch == '\t'
                ? column + 4 - ((column - 1) % 4)
                : column + 1;
        }

        return startIndex >= 0 && endIndex >= startIndex;
    }

    private static bool LooksLikeSetextHeadingUnderline(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) > 3) return false;

        var trimmed = line.Trim();
        if (trimmed.Length == 0 || trimmed[0] != '-') return false;
        for (int i = 0; i < trimmed.Length; i++) {
            if (trimmed[i] != '-') return false;
        }

        return true;
    }
}
