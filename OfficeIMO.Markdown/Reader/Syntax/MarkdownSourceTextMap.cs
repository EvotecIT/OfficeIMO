namespace OfficeIMO.Markdown;

internal sealed class MarkdownSourceTextMap {
    private readonly string _text;
    private readonly int[] _lineStarts;

    internal MarkdownSourceTextMap(string text) {
        _text = text ?? string.Empty;
        _lineStarts = BuildLineStarts(_text);
    }

    internal MarkdownSourceSpan CreateLineSpan(int startLine, int endLine) {
        startLine = Math.Max(1, startLine);
        endLine = Math.Max(startLine, endLine);

        var startColumn = 1;
        var endColumn = GetLineLength(endLine);
        if (endColumn < 1) {
            endColumn = 1;
        }

        return CreateSpan(startLine, startColumn, endLine, endColumn);
    }

    internal MarkdownSourceSpan CreateSpan(int startLine, int startColumn, int endLine, int endColumn) {
        startLine = Math.Max(1, startLine);
        endLine = Math.Max(startLine, endLine);

        var normalizedStartColumn = NormalizeColumn(startLine, startColumn);
        var normalizedEndColumn = NormalizeColumn(endLine, endColumn);
        var startOffset = GetOffset(startLine, normalizedStartColumn);
        var endOffset = GetOffset(endLine, normalizedEndColumn);

        return new MarkdownSourceSpan(
            startLine,
            normalizedStartColumn,
            endLine,
            normalizedEndColumn,
            startOffset,
            endOffset);
    }

    internal MarkdownSourcePoint CreatePoint(int line, int column) {
        var normalizedLine = Math.Max(1, line);
        var normalizedColumn = NormalizeColumn(normalizedLine, column);
        return new MarkdownSourcePoint(normalizedLine, normalizedColumn, GetOffset(normalizedLine, normalizedColumn));
    }

    private int NormalizeColumn(int line, int column) {
        var length = GetLineLength(line);
        if (length <= 0) {
            return 1;
        }

        if (column < 1) {
            return 1;
        }

        if (column > length) {
            return length;
        }

        return column;
    }

    private int GetLineLength(int line) {
        if (line < 1 || line > _lineStarts.Length) {
            return 1;
        }

        var lineStart = _lineStarts[line - 1];
        var lineEndExclusive = line < _lineStarts.Length ? _lineStarts[line] - 1 : _text.Length;
        while (lineEndExclusive > lineStart && _text[lineEndExclusive - 1] == '\n') {
            lineEndExclusive--;
        }

        return Math.Max(1, lineEndExclusive - lineStart);
    }

    private int GetOffset(int line, int column) {
        if (_text.Length == 0 || line < 1 || line > _lineStarts.Length) {
            return 0;
        }

        var lineStart = _lineStarts[line - 1];
        return Math.Min(_text.Length - 1, lineStart + Math.Max(0, column - 1));
    }

    private static int[] BuildLineStarts(string text) {
        if (string.IsNullOrEmpty(text)) {
            return new[] { 0 };
        }

        var starts = new List<int> {
            0
        };
        for (var i = 0; i < text.Length; i++) {
            if (text[i] == '\n' && i + 1 < text.Length) {
                starts.Add(i + 1);
            }
        }

        return starts.ToArray();
    }
}
