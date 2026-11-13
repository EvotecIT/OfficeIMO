using System.Text;

namespace OfficeIMO.Markdown;

internal sealed class MarkdownInlineSourceMap {
    private readonly MarkdownSourcePoint?[] _points;
    private readonly MarkdownSourceSpan?[] _tokenSpans;
    private readonly string?[] _tokenLiterals;
    private readonly string? _singleLineText;
    private readonly int _singleLineTextStart;
    private readonly int _singleLineTextLength;
    private readonly int _singleLine;
    private readonly int _singleLineStartColumn;
    private readonly int _singleLineStartOffset;
    // Column at every index of _singleLineText, kept only when the line has tabs. Without tabs a
    // column is the start column plus the distance, so spans no longer rescan the line prefix.
    private readonly int[]? _singleLineTabColumns;

    internal MarkdownInlineSourceMap(
        MarkdownSourceTextMap sourceTextMap,
        string text,
        int absoluteLine,
        int startColumn) {
        _points = Array.Empty<MarkdownSourcePoint?>();
        _tokenSpans = Array.Empty<MarkdownSourceSpan?>();
        _tokenLiterals = Array.Empty<string?>();
        ArgumentNullException.ThrowIfNull(sourceTextMap);
        var startPoint = sourceTextMap.CreatePoint(absoluteLine, startColumn);
        _singleLineText = text ?? string.Empty;
        _singleLineTextLength = _singleLineText.Length;
        _singleLine = startPoint.Line;
        _singleLineStartColumn = startPoint.Column;
        _singleLineStartOffset = startPoint.Offset;
        _singleLineTabColumns = BuildTabColumns(_singleLineText, _singleLineStartColumn);
    }

    private MarkdownInlineSourceMap(
        string text,
        int textStart,
        int textLength,
        int absoluteLine,
        int startColumn,
        int startOffset,
        int[]? tabColumns) {
        _points = Array.Empty<MarkdownSourcePoint?>();
        _tokenSpans = Array.Empty<MarkdownSourceSpan?>();
        _tokenLiterals = Array.Empty<string?>();
        _singleLineText = text ?? string.Empty;
        _singleLineTextStart = textStart;
        _singleLineTextLength = textLength;
        _singleLine = absoluteLine;
        _singleLineStartColumn = startColumn;
        _singleLineStartOffset = startOffset;
        _singleLineTabColumns = tabColumns;
    }

    internal MarkdownInlineSourceMap(MarkdownSourcePoint?[] points) {
        _points = points ?? Array.Empty<MarkdownSourcePoint?>();
        _tokenSpans = Array.Empty<MarkdownSourceSpan?>();
        _tokenLiterals = Array.Empty<string?>();
    }

    internal MarkdownInlineSourceMap(
        MarkdownSourcePoint?[] points,
        MarkdownSourceSpan?[]? tokenSpans,
        string?[]? tokenLiterals) {
        _points = points ?? Array.Empty<MarkdownSourcePoint?>();
        _tokenSpans = tokenSpans ?? Array.Empty<MarkdownSourceSpan?>();
        _tokenLiterals = tokenLiterals ?? Array.Empty<string?>();
    }

    internal int Length => _singleLineText != null ? _singleLineTextLength : Math.Max(_points.Length, _tokenLiterals.Length);

    internal MarkdownSourceSpan? GetSpan(int startIndex, int length) {
        if (_singleLineText != null) {
            if (length <= 0 || startIndex < 0 || startIndex >= _singleLineTextLength) {
                return null;
            }

            var endExclusive = Math.Min(_singleLineTextLength, startIndex + length);
            var sourceStart = _singleLineTextStart + startIndex;
            var sourceEndExclusive = _singleLineTextStart + endExclusive;
            var startColumn = GetSingleLineColumn(sourceStart);
            var endColumnExclusive = GetSingleLineColumn(sourceEndExclusive);
            return new MarkdownSourceSpan(
                _singleLine,
                startColumn,
                _singleLine,
                Math.Max(startColumn, endColumnExclusive - 1),
                _singleLineStartOffset + startIndex,
                _singleLineStartOffset + endExclusive - 1);
        }

        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return null;
        }

        if (length == 1 && startIndex < _tokenSpans.Length && _tokenSpans[startIndex].HasValue) {
            return _tokenSpans[startIndex];
        }

        var endIndex = Math.Min(_points.Length - 1, startIndex + length - 1);
        // Only the first and last mapped positions matter. Walking inward from both ends keeps a
        // wide span (an emphasis around a long paragraph) from costing its full length.
        MarkdownSourcePoint? start = null;
        MarkdownSourcePoint? end = null;

        for (var i = startIndex; i <= endIndex && !start.HasValue; i++) {
            if (i < _tokenSpans.Length && _tokenSpans[i].HasValue) {
                var span = _tokenSpans[i]!.Value;
                start = new MarkdownSourcePoint(span.StartLine, span.StartColumn ?? 1, span.StartOffset ?? 0);
            } else {
                start = _points[i];
            }
        }

        for (var i = endIndex; i >= startIndex && start.HasValue && !end.HasValue; i--) {
            if (i < _tokenSpans.Length && _tokenSpans[i].HasValue) {
                var span = _tokenSpans[i]!.Value;
                end = new MarkdownSourcePoint(span.EndLine, span.EndColumn ?? 1, span.EndOffset ?? 0);
            } else {
                end = _points[i];
            }
        }

        if (!start.HasValue || !end.HasValue) {
            return null;
        }

        return new MarkdownSourceSpan(
            start.Value.Line,
            start.Value.Column,
            end.Value.Line,
            end.Value.Column,
            start.Value.Offset,
            end.Value.Offset);
    }

    internal string? GetTokenLiteral(int startIndex, int length) {
        if (_singleLineText != null) {
            return null;
        }

        if (length != 1 || startIndex < 0 || startIndex >= _tokenLiterals.Length) {
            return null;
        }

        return _tokenLiterals[startIndex];
    }

    internal bool ContainsSourceLineBreak(int startIndex, int length) {
        if (_singleLineText != null) {
            return false;
        }

        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return false;
        }

        var endExclusive = Math.Min(_points.Length, startIndex + length);
        for (var i = startIndex; i < endExclusive; i++) {
            if (i + 1 >= _points.Length) {
                break;
            }

            var current = _points[i];
            var next = _points[i + 1];
            if (current.HasValue && next.HasValue && current.Value.Line != next.Value.Line) {
                return true;
            }
        }

        return false;
    }

    internal string RestoreSourceLineBreaks(string text, int startIndex, int length) {
        if (string.IsNullOrEmpty(text) || length <= 0 || startIndex < 0 || startIndex >= text.Length) {
            return string.Empty;
        }

        if (_singleLineText != null) {
            return text.Substring(startIndex, Math.Min(length, text.Length - startIndex));
        }

        var endExclusive = Math.Min(text.Length, startIndex + length);
        var sb = new StringBuilder(endExclusive - startIndex);
        for (var i = startIndex; i < endExclusive; i++) {
            var ch = text[i];
            if (ch == ' ' && i + 1 < endExclusive && HasSourceLineTransitionAfter(i)) {
                sb.Append('\n');
                continue;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }

    private bool HasSourceLineTransitionAfter(int index) {
        if (index < 0 || index + 1 >= _points.Length) {
            return false;
        }

        var current = _points[index];
        var next = _points[index + 1];
        return current.HasValue && next.HasValue && current.Value.Line != next.Value.Line;
    }

    internal MarkdownInlineSourceMap Slice(int startIndex, int length) {
        if (_singleLineText != null) {
            if (length <= 0 || startIndex < 0 || startIndex >= _singleLineTextLength) {
                return new MarkdownInlineSourceMap(Array.Empty<MarkdownSourcePoint?>());
            }

            var directLength = Math.Min(length, _singleLineTextLength - startIndex);
            var directStart = _singleLineTextStart + startIndex;
            return new MarkdownInlineSourceMap(
                _singleLineText,
                directStart,
                directLength,
                _singleLine,
                GetSingleLineColumn(directStart),
                _singleLineStartOffset + startIndex,
                _singleLineTabColumns);
        }

        if (length <= 0 || startIndex < 0 || startIndex >= Length) {
            return new MarkdownInlineSourceMap(Array.Empty<MarkdownSourcePoint?>());
        }

        var actualLength = Math.Min(length, Length - startIndex);
        var slice = Array.Empty<MarkdownSourcePoint?>();
        if (_points.Length > startIndex) {
            slice = new MarkdownSourcePoint?[actualLength];
            Array.Copy(_points, startIndex, slice, 0, Math.Min(actualLength, _points.Length - startIndex));
        }

        var tokenSpanSlice = Array.Empty<MarkdownSourceSpan?>();
        if (_tokenSpans.Length > startIndex) {
            tokenSpanSlice = new MarkdownSourceSpan?[actualLength];
            Array.Copy(_tokenSpans, startIndex, tokenSpanSlice, 0, Math.Min(actualLength, _tokenSpans.Length - startIndex));
        }

        var tokenLiteralSlice = Array.Empty<string?>();
        if (_tokenLiterals.Length > startIndex) {
            tokenLiteralSlice = new string?[actualLength];
            Array.Copy(_tokenLiterals, startIndex, tokenLiteralSlice, 0, Math.Min(actualLength, _tokenLiterals.Length - startIndex));
        }

        return new MarkdownInlineSourceMap(slice, tokenSpanSlice, tokenLiteralSlice);
    }

    private int GetSingleLineColumn(int sourceIndex) =>
        _singleLineTabColumns != null
            ? _singleLineTabColumns[sourceIndex]
            : _singleLineStartColumn + (sourceIndex - _singleLineTextStart);

    private static int[]? BuildTabColumns(string text, int startColumn) {
        if (text.IndexOf('\t') < 0) {
            return null;
        }

        var columns = new int[text.Length + 1];
        columns[0] = startColumn;
        for (var index = 0; index < text.Length; index++) {
            columns[index + 1] = MarkdownSourceColumns.AdvanceColumn(columns[index], text[index]);
        }

        return columns;
    }
}

internal readonly struct MarkdownSourcePoint {
    internal MarkdownSourcePoint(int line, int column, int offset) {
        Line = line;
        Column = column;
        Offset = offset;
    }

    internal int Line { get; }
    internal int Column { get; }
    internal int Offset { get; }
}
