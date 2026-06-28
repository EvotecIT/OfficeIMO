namespace OfficeIMO.Markdown;

internal static class HtmlBlockSourceSpanHelpers {
    internal static MarkdownSourceSpan? GetSourceSpan(string text, MarkdownSourceSpan? blockSpan, int startIndex, int endIndex) {
        if (!blockSpan.HasValue || !blockSpan.Value.StartColumn.HasValue || startIndex < 0 || endIndex < startIndex || endIndex >= text.Length) {
            return null;
        }

        var start = GetPoint(text, blockSpan.Value, startIndex);
        var end = GetPoint(text, blockSpan.Value, endIndex);
        return new MarkdownSourceSpan(start.Line, start.Column, end.Line, end.Column);
    }

    private static (int Line, int Column) GetPoint(string text, MarkdownSourceSpan blockSpan, int index) {
        var line = blockSpan.StartLine;
        var column = blockSpan.StartColumn ?? 1;

        for (var i = 0; i < index && i < text.Length; i++) {
            if (text[i] == '\n') {
                line++;
                column = 1;
            } else {
                column++;
            }
        }

        return (line, column);
    }
}
