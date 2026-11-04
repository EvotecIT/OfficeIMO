using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static int GetListContinuationIndent(string line, MarkdownLoadOptions? options = null) {
        if (string.IsNullOrEmpty(line)) return 0;
        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out int orderedContentStartIndex, out _, out _)) {
            if (IsWhiteSpaceFrom(line, orderedContentStartIndex)
                && TryGetOrderedListMarkerWidth(line, orderedLeadingSpaces, options, out int orderedMarkerWidth)) {
                return orderedLeadingSpaces + orderedMarkerWidth + 1;
            }

            return orderedContentStartIndex;
        }

        if (TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out int unorderedContentStartIndex)) {
            if (IsWhiteSpaceFrom(line, unorderedContentStartIndex)) {
                return unorderedLeadingSpaces + 2;
            }

            return unorderedContentStartIndex;
        }

        int spaces = CountLeadingSpaces(line);
        return spaces + 2;
    }

    private static bool IsWhiteSpaceFrom(string value, int start) {
        if (string.IsNullOrEmpty(value) || start >= value.Length) {
            return true;
        }

        for (int i = Math.Max(0, start); i < value.Length; i++) {
            if (!char.IsWhiteSpace(value[i])) {
                return false;
            }
        }

        return true;
    }

    private static int GetRelativeListItemLevel(List<int>? continuationIndentsByLevel, string line) {
        if (continuationIndentsByLevel == null || continuationIndentsByLevel.Count == 0 || string.IsNullOrEmpty(line)) {
            return 0;
        }

        int indentColumns = CountLeadingIndentColumns(line);
        for (int level = continuationIndentsByLevel.Count - 1; level >= 0; level--) {
            if (indentColumns >= continuationIndentsByLevel[level]) {
                return level + 1;
            }
        }

        return 0;
    }

    private static void TrackListItemContinuationIndent(List<int> continuationIndentsByLevel, int level, int continuationIndent) {
        if (continuationIndentsByLevel == null) {
            return;
        }

        while (continuationIndentsByLevel.Count > level) {
            continuationIndentsByLevel.RemoveAt(continuationIndentsByLevel.Count - 1);
        }

        if (continuationIndentsByLevel.Count == level) {
            continuationIndentsByLevel.Add(continuationIndent);
            return;
        }

        continuationIndentsByLevel[level] = continuationIndent;
    }

    private static int GetTaskMarkerConsumedColumns(string content) {
        if (string.IsNullOrEmpty(content)) return 0;
        return TryGetTaskListMarkerContentStartIndex(content, out int contentStartIndex)
            ? contentStartIndex
            : 0;
    }

    private static void SetListItemMarkerSourceSpans(ListItem item, string line, int lineIndex, bool isTask, MarkdownLoadOptions options, MarkdownReaderState state) {
        if (item == null) {
            return;
        }

        item.MarkerText = TryGetListMarkerText(line, options);
        if (isTask) {
            item.TaskMarkerText = TryGetTaskMarkerText(line, options);
        } else {
            item.TaskMarkerText = null;
        }

        if (!state.CaptureSyntaxTree) {
            return;
        }

        int absoluteLineNumber = state.SourceLineOffset + lineIndex + 1;
        item.MarkerSourceSpan = TryCreateListMarkerSourceSpan(line, absoluteLineNumber, options, state);
        if (isTask) {
            item.TaskMarkerSourceSpan = TryCreateTaskMarkerSourceSpan(line, absoluteLineNumber, options, state);
        }
    }

    private static string? TryGetListMarkerText(string line, MarkdownLoadOptions? options = null) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out _, out _, out _)) {
            int delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, orderedLeadingSpaces);
            return delimiterIndex < line.Length
                ? line.Substring(orderedLeadingSpaces, delimiterIndex - orderedLeadingSpaces + 1)
                : null;
        }

        if (!TryGetUnorderedListMarkerInfo(line, out _, out _, out char marker)) {
            return null;
        }

        return marker switch {
            '-' => "-",
            '*' => "*",
            '+' => "+",
            _ => null
        };
    }

    private static string? TryGetTaskMarkerText(string line, MarkdownLoadOptions? options = null) {
        if (string.IsNullOrEmpty(line)
            || !TryGetRawListItemContentAfterMarker(line, out string content, options)
            || !TryGetTaskListMarkerContentStartIndex(content, out _)) {
            return null;
        }

        return content.Length >= 3 ? content.Substring(0, 3) : null;
    }

    private static MarkdownSourceSpan? TryCreateListMarkerSourceSpan(string line, int absoluteLineNumber, MarkdownLoadOptions? options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (TryGetOrderedListMarkerInfo(line, options, out int orderedLeadingSpaces, out _, out _, out _, out _)) {
            int delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, orderedLeadingSpaces);
            if (delimiterIndex < line.Length) {
                return CreateSpan(state, absoluteLineNumber, orderedLeadingSpaces + 1, absoluteLineNumber, delimiterIndex + 1);
            }
        }

        if (TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out _, out _)) {
            int markerColumn = unorderedLeadingSpaces + 1;
            return CreateSpan(state, absoluteLineNumber, markerColumn, absoluteLineNumber, markerColumn);
        }

        return null;
    }

    private static MarkdownSourceSpan? TryCreateTaskMarkerSourceSpan(string line, int absoluteLineNumber, MarkdownLoadOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(line)) {
            return null;
        }

        if (!TryGetRawListItemContentAfterMarker(line, out string content, options)
            || !TryGetTaskListMarkerContentStartIndex(content, out _)) {
            return null;
        }

        int startColumn = GetListContinuationIndent(line, options) + 1;
        return CreateSpan(state, absoluteLineNumber, startColumn, absoluteLineNumber, startColumn + 2);
    }

    private static bool TryStripTaskListMarker(string content, out bool done, out string stripped) {
        done = false;
        stripped = content ?? string.Empty;
        if (!TryGetTaskListMarkerContentStartIndex(stripped, out int contentStartIndex)) {
            return false;
        }

        done = stripped[1] == 'x' || stripped[1] == 'X';
        stripped = contentStartIndex >= stripped.Length ? string.Empty : stripped.Substring(contentStartIndex);
        return true;
    }

    private static bool TryGetTaskListMarkerContentStartIndex(string content, out int contentStartIndex) {
        contentStartIndex = 0;
        if (string.IsNullOrEmpty(content) || content.Length < 3) {
            return false;
        }

        if (content[0] != '[' || content[2] != ']') {
            return false;
        }

        if (content[1] != ' ' && content[1] != 'x' && content[1] != 'X') {
            return false;
        }

        if (content.Length == 3) {
            contentStartIndex = 3;
            return true;
        }

        if (!char.IsWhiteSpace(content[3])) {
            return false;
        }

        contentStartIndex = 4;
        while (contentStartIndex < content.Length && char.IsWhiteSpace(content[contentStartIndex])) {
            contentStartIndex++;
        }

        return true;
    }

    private static bool TryGetTaskListMarkerContentStartIndex(string content, int start, out int contentStartIndex) {
        contentStartIndex = start;
        if (string.IsNullOrEmpty(content) || start < 0 || start + 2 >= content.Length) {
            return false;
        }

        if (content[start] != '[' || content[start + 2] != ']') {
            return false;
        }

        char marker = content[start + 1];
        if (marker != ' ' && marker != 'x' && marker != 'X') {
            return false;
        }

        int position = start + 3;
        if (position == content.Length) {
            contentStartIndex = position;
            return true;
        }

        if (!char.IsWhiteSpace(content[position])) {
            return false;
        }

        position++;
        while (position < content.Length && char.IsWhiteSpace(content[position])) {
            position++;
        }

        contentStartIndex = position;
        return true;
    }

    private static bool TryGetRawListItemContentAfterMarker(string line, out string content, MarkdownLoadOptions? options = null) {
        content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        if (options?.OrderedLists != false &&
            TryGetOrderedListMarkerInfo(line, options, out var orderedLeadingSpaces, out _, out int orderedContentStartIndex, out _, out _)) {
            if (options?.StrictListIndentation == true && orderedLeadingSpaces > 3) {
                return false;
            }

            content = line.Substring(orderedContentStartIndex);
            return true;
        }

        if (options?.UnorderedLists != false &&
            TryGetUnorderedListMarkerInfo(line, out var unorderedLeadingSpaces, out int unorderedContentStartIndex)) {
            if (options?.StrictListIndentation == true && unorderedLeadingSpaces > 3) {
                return false;
            }

            content = line.Substring(unorderedContentStartIndex);
            return true;
        }

        return false;
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex) {
        return TryGetOrderedListMarkerInfo(line, null, out leadingSpaces, out number, out contentStartIndex, out _, out _);
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex, out char delimiter) {
        return TryGetOrderedListMarkerInfo(line, null, out leadingSpaces, out number, out contentStartIndex, out delimiter, out _);
    }

    private static bool TryGetOrderedListMarkerInfo(
        string line,
        MarkdownLoadOptions? options,
        out int leadingSpaces,
        out int number,
        out int contentStartIndex,
        out char delimiter,
        out MarkdownOrderedListMarkerStyle markerStyle) {
        leadingSpaces = 0;
        number = 0;
        contentStartIndex = 0;
        delimiter = '\0';
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;

        int markerStart = leadingSpaces;
        int markerEnd = markerStart;
        while (markerEnd < line.Length && char.IsDigit(line[markerEnd])) markerEnd++;
        if (markerEnd > markerStart) {
            if (markerEnd - markerStart > 9) return false;
            if (markerEnd >= line.Length || (line[markerEnd] != '.' && line[markerEnd] != ')')) return false;
            delimiter = line[markerEnd];
            if (!TryGetListContentStartIndex(line, markerEnd, out contentStartIndex)) return false;
            if (!int.TryParse(line.Substring(markerStart, markerEnd - markerStart), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
            return true;
        }

        if (options?.ListExtras != true) return false;

        while (markerEnd < line.Length && IsAsciiLetter(line[markerEnd])) markerEnd++;
        if (markerEnd == markerStart) return false;
        if (markerEnd >= line.Length || (line[markerEnd] != '.' && line[markerEnd] != ')')) return false;
        delimiter = line[markerEnd];
        if (!TryGetListContentStartIndex(line, markerEnd, out contentStartIndex)) return false;

        var marker = line.Substring(markerStart, markerEnd - markerStart);
        if (TryParseRomanListExtraMarker(marker, out number, out markerStyle)) {
            return true;
        }

        if (marker.Length == 1) {
            var ch = marker[0];
            markerStyle = char.IsUpper(ch) ? MarkdownOrderedListMarkerStyle.UpperAlpha : MarkdownOrderedListMarkerStyle.LowerAlpha;
            number = char.ToUpperInvariant(ch) - 'A' + 1;
            return number >= 1 && number <= 26;
        }

        number = 0;
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        return false;
    }

    private static int GetOrderedListMarkerDelimiterIndex(string line, int leadingSpaces) {
        var delimiterIndex = leadingSpaces;
        while (delimiterIndex < line.Length
               && (char.IsDigit(line[delimiterIndex]) || IsAsciiLetter(line[delimiterIndex]))) {
            delimiterIndex++;
        }

        return delimiterIndex;
    }

    private static bool TryParseRomanListExtraMarker(string marker, out int number, out MarkdownOrderedListMarkerStyle markerStyle) {
        number = 0;
        markerStyle = MarkdownOrderedListMarkerStyle.Decimal;
        if (string.IsNullOrEmpty(marker)) {
            return false;
        }

        var upper = char.IsUpper(marker[0]);
        var lower = marker.ToLowerInvariant();
        for (int i = 0; i < lower.Length; i++) {
            var ch = lower[i];
            if (ch != 'i' && ch != 'v' && ch != 'x') {
                return false;
            }
        }

        for (int value = 1; value <= 39; value++) {
            if (string.Equals(lower, FormatListExtraRoman(value), StringComparison.Ordinal)) {
                number = value;
                markerStyle = upper
                    ? MarkdownOrderedListMarkerStyle.UpperRoman
                    : MarkdownOrderedListMarkerStyle.LowerRoman;
                return true;
            }
        }

        return false;
    }

    private static string FormatListExtraRoman(int value) {
        var tens = value / 10;
        var ones = value % 10;
        return new string('x', tens) + (ones switch {
            0 => string.Empty,
            1 => "i",
            2 => "ii",
            3 => "iii",
            4 => "iv",
            5 => "v",
            6 => "vi",
            7 => "vii",
            8 => "viii",
            9 => "ix",
            _ => string.Empty
        });
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex) {
        return TryGetUnorderedListMarkerInfo(line, out leadingSpaces, out contentStartIndex, out _);
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex, out char marker) {
        leadingSpaces = 0;
        contentStartIndex = 0;
        marker = '\0';
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        marker = line[leadingSpaces];
        if (marker != '-' && marker != '*' && marker != '+') return false;
        return TryGetListContentStartIndex(line, leadingSpaces, out contentStartIndex);
    }

    private static bool TryGetListContentStartIndex(string line, int markerIndex, out int contentStartIndex) {
        contentStartIndex = 0;
        int paddingStart = markerIndex + 1;
        if (paddingStart >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        int paddingColumns = 0;
        int cursor = paddingStart;
        while (cursor < line.Length) {
            char ch = line[cursor];
            if (ch == ' ' && paddingColumns < 4) {
                paddingColumns++;
                cursor++;
                continue;
            }

            if (ch == '\t' && paddingColumns == 0) {
                contentStartIndex = cursor + 1;
                return true;
            }

            break;
        }

        if (cursor >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        if (paddingColumns == 0) return false;
        contentStartIndex = cursor;
        return true;
    }

    private static bool TryGetIndentedCodeListLead(string line, MarkdownLoadOptions? options, out int continuationIndent, out string content, out int startColumn) {
        continuationIndent = 0;
        content = string.Empty;
        startColumn = 1;
        if (string.IsNullOrEmpty(line)) return false;

        int leadingSpaces = 0;
        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        int markerWidth;
        if (TryGetOrderedListMarkerWidth(line, leadingSpaces, options, out markerWidth)) {
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces + markerWidth - 1)) return false;
        } else {
            char marker = line[leadingSpaces];
            if (marker != '-' && marker != '*' && marker != '+') return false;
            markerWidth = 1;
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces)) return false;
        }

        continuationIndent = leadingSpaces + markerWidth + 1;
        if (continuationIndent >= line.Length) return false;

        content = ExpandWhitespaceAfterContainerMarker(line, leadingSpaces + markerWidth).Substring(continuationIndent);
        startColumn = continuationIndent + 1;
        return CountLeadingIndentColumns(content) >= 4;
    }

    private static bool TryGetOrderedListMarkerWidth(string line, int leadingSpaces, out int markerWidth) {
        return TryGetOrderedListMarkerWidth(line, leadingSpaces, null, out markerWidth);
    }

    private static bool TryGetOrderedListMarkerWidth(string line, int leadingSpaces, MarkdownLoadOptions? options, out int markerWidth) {
        markerWidth = 0;
        if (string.IsNullOrEmpty(line) || leadingSpaces >= line.Length) return false;

        if (!TryGetOrderedListMarkerInfo(line, options, out _, out _, out _, out _, out _)) {
            return false;
        }

        var delimiterIndex = GetOrderedListMarkerDelimiterIndex(line, leadingSpaces);
        markerWidth = delimiterIndex - leadingSpaces + 1;
        return true;
    }

    private static bool HasIndentedCodePaddingAfterMarker(string line, int markerEndIndex) {
        line = ExpandWhitespaceAfterContainerMarker(line, markerEndIndex + 1);
        int paddingStart = markerEndIndex + 1;
        if (paddingStart >= line.Length || line[paddingStart] != ' ') return false;

        int spaces = 0;
        int cursor = paddingStart;
        while (cursor < line.Length && line[cursor] == ' ') {
            spaces++;
            cursor++;
        }

        return spaces >= 5;
    }

    private static int GetListLeadContentStartColumn(string line, MarkdownLoadOptions? options = null, bool stripTaskMarker = false) {
        int startColumn = GetListContinuationIndent(line, options) + 1;
        if (!stripTaskMarker) return startColumn;

        return TryGetRawListItemContentAfterMarker(line, out string content, options)
            ? startColumn + GetTaskMarkerConsumedColumns(content)
            : startColumn;
    }
}
