namespace OfficeIMO.Markdown;

/// <summary>
/// Shared helpers for Markdown's 1-based, tab-expanded source columns.
/// </summary>
internal static class MarkdownSourceColumns {
    internal static int GetColumnWidth(int currentColumns, char value) =>
        value == '\t'
            ? 4 - (currentColumns % 4)
            : 1;

    internal static int AdvanceColumn(int column, char value) =>
        column + GetColumnWidth(Math.Max(0, column - 1), value);

    internal static int ResolveVisualColumnOffset(string sourceText, int lineStart, int columnNumber) {
        sourceText ??= string.Empty;
        var lineEndExclusive = sourceText.Length;
        for (var index = Math.Max(0, lineStart); index < sourceText.Length; index++) {
            if (IsLineBreakStart(sourceText, index, out _)) {
                lineEndExclusive = index;
                break;
            }
        }

        return ResolveVisualColumnOffset(sourceText, lineStart, lineEndExclusive, columnNumber);
    }

    internal static int ResolveVisualColumnOffset(string sourceText, int lineStart, int lineEndExclusive, int columnNumber) {
        sourceText ??= string.Empty;
        if (sourceText.Length == 0) {
            return 0;
        }

        var normalizedLineStart = Math.Max(0, Math.Min(sourceText.Length - 1, lineStart));
        var normalizedLineEndExclusive = Math.Max(normalizedLineStart, Math.Min(sourceText.Length, lineEndExclusive));
        var normalizedColumn = Math.Max(1, columnNumber);
        var columns = 0;
        var lastCharacterOffset = normalizedLineStart;
        for (var index = normalizedLineStart; index < normalizedLineEndExclusive; index++) {
            lastCharacterOffset = index;
            columns += GetColumnWidth(columns, sourceText[index]);

            if (normalizedColumn <= columns) {
                return index;
            }
        }

        return Math.Min(sourceText.Length - 1, lastCharacterOffset);
    }

    internal static int GetColumnNumber(string sourceText, int offset) {
        sourceText ??= string.Empty;
        var column = 1;
        var lineStartOffset = 0;
        var boundedOffset = Math.Max(0, Math.Min(sourceText.Length, offset));
        for (var i = boundedOffset - 1; i >= 0; i--) {
            if (IsLineBreakStart(sourceText, i, out _)) {
                lineStartOffset = i + 1;
                break;
            }
        }

        for (var i = lineStartOffset; i < boundedOffset && i < sourceText.Length; i++) {
            column = AdvanceColumn(column, sourceText[i]);
        }

        return column;
    }

    internal static int GetEndColumnNumber(string sourceText, int startOffset, int length, int startColumn) {
        sourceText ??= string.Empty;
        if (sourceText.Length == 0 || length <= 0) {
            return startColumn;
        }

        var normalizedStartOffset = Math.Max(0, Math.Min(sourceText.Length - 1, startOffset));
        var endExclusive = Math.Min(sourceText.Length, normalizedStartOffset + length);
        var column = startColumn;
        for (var i = normalizedStartOffset; i < endExclusive; i++) {
            if (i > normalizedStartOffset) {
                column = AdvanceColumn(column, sourceText[i - 1]);
            }
        }

        return sourceText[endExclusive - 1] == '\t'
            ? AdvanceColumn(column, '\t') - 1
            : column;
    }

    internal static bool IsLineBreakStart(string sourceText, int offset, out int length) {
        if (offset >= 0 && offset < sourceText.Length && sourceText[offset] == '\r') {
            length = offset + 1 < sourceText.Length && sourceText[offset + 1] == '\n'
                ? 2
                : 1;
            return true;
        }

        if (offset >= 0 && offset < sourceText.Length && sourceText[offset] == '\n') {
            length = 1;
            return true;
        }

        length = 0;
        return false;
    }
}
