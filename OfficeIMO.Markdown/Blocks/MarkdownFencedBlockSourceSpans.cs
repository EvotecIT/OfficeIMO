namespace OfficeIMO.Markdown;

internal static class MarkdownFencedBlockSourceSpans {
    internal static MarkdownSourceSpan? GetOpeningFenceSpan(MarkdownSourceSpan? span, bool isFenced, int fenceIndentColumns, int fenceLength) {
        if (!isFenced || !span.HasValue) {
            return null;
        }

        var value = span.Value;
        if (!value.StartColumn.HasValue) {
            return new MarkdownSourceSpan(value.StartLine, value.StartLine);
        }

        var startColumn = value.StartColumn.Value + Math.Max(0, fenceIndentColumns);
        var endColumn = startColumn + Math.Max(3, fenceLength) - 1;
        int? startOffset = null;
        int? endOffset = null;
        if (value.StartOffset.HasValue) {
            startOffset = value.StartOffset.Value + Math.Max(0, fenceIndentColumns);
            endOffset = startOffset.Value + Math.Max(3, fenceLength) - 1;
        }

        return new MarkdownSourceSpan(value.StartLine, startColumn, value.StartLine, endColumn, startOffset, endOffset);
    }

    internal static MarkdownSourceSpan? GetInfoSpan(
        MarkdownSourceSpan? span,
        bool isFenced,
        string? infoString,
        int fenceIndentColumns,
        int fenceLength,
        int infoPaddingColumns) {
        if (!isFenced || !span.HasValue || string.IsNullOrEmpty(infoString)) {
            return null;
        }

        var value = span.Value;
        if (!value.StartColumn.HasValue) {
            return new MarkdownSourceSpan(value.StartLine, value.StartLine);
        }

        var startColumn = value.StartColumn.Value
            + Math.Max(0, fenceIndentColumns)
            + Math.Max(3, fenceLength)
            + Math.Max(0, infoPaddingColumns);
        var endColumn = startColumn + infoString!.Length - 1;
        int? startOffset = null;
        int? endOffset = null;
        if (value.StartOffset.HasValue) {
            startOffset = value.StartOffset.Value
                + Math.Max(0, fenceIndentColumns)
                + Math.Max(3, fenceLength)
                + Math.Max(0, infoPaddingColumns);
            endOffset = startOffset.Value + infoString.Length - 1;
        }

        return new MarkdownSourceSpan(value.StartLine, startColumn, value.StartLine, endColumn, startOffset, endOffset);
    }

    internal static MarkdownSourceSpan? GetClosingFenceSpan(
        MarkdownSourceSpan? span,
        bool isFenced,
        string? content,
        bool hasClosingFence,
        int closingFenceIndentColumns,
        int closingFenceLength) {
        if (!isFenced || !hasClosingFence || !span.HasValue) {
            return null;
        }

        var value = span.Value;
        var line = value.StartLine + CountContentLines(content) + 1;
        if (line > value.EndLine) {
            return null;
        }

        if (!value.StartColumn.HasValue) {
            return new MarkdownSourceSpan(line, line);
        }

        var startColumn = value.StartColumn.Value + Math.Max(0, closingFenceIndentColumns);
        var endColumn = startColumn + Math.Max(3, closingFenceLength) - 1;
        return new MarkdownSourceSpan(line, startColumn, line, endColumn);
    }

    internal static MarkdownSourceSpan? GetContentSpan(MarkdownSourceSpan? span, bool isFenced, string? content) {
        if (!span.HasValue) {
            return null;
        }

        var value = span.Value;
        if (!isFenced) {
            return value;
        }

        var startLine = value.StartLine + 1;
        var endLine = value.StartLine + CountContentLines(content);
        if (endLine < startLine || endLine > value.EndLine) {
            return null;
        }

        if (!value.StartColumn.HasValue) {
            return new MarkdownSourceSpan(startLine, endLine);
        }

        return new MarkdownSourceSpan(startLine, 1, endLine, GetLastContentLineColumn(content));
    }

    private static int CountContentLines(string? content) {
        if (string.IsNullOrEmpty(content)) {
            return 0;
        }

        var count = 1;
        for (var i = 0; i < content!.Length; i++) {
            if (content[i] == '\n') {
                count++;
            }
        }

        return count;
    }

    private static int GetLastContentLineColumn(string? content) {
        if (string.IsNullOrEmpty(content)) {
            return 1;
        }

        var lastBreak = content!.LastIndexOf('\n');
        var lastLine = lastBreak >= 0 ? content.Substring(lastBreak + 1) : content;
        return Math.Max(1, lastLine.Length);
    }
}
