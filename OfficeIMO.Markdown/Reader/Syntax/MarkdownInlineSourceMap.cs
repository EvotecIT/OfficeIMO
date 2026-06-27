namespace OfficeIMO.Markdown;

internal sealed class MarkdownInlineSourceMap {
    private readonly MarkdownSourcePoint?[] _points;
    private readonly MarkdownSourceSpan?[] _tokenSpans;
    private readonly string?[] _tokenLiterals;

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

    internal int Length => _points.Length;

    internal MarkdownSourceSpan? GetSpan(int startIndex, int length) {
        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return null;
        }

        if (length == 1 && startIndex < _tokenSpans.Length && _tokenSpans[startIndex].HasValue) {
            return _tokenSpans[startIndex];
        }

        var endIndex = Math.Min(_points.Length - 1, startIndex + length - 1);
        MarkdownSourcePoint? start = null;
        MarkdownSourcePoint? end = null;

        for (var i = startIndex; i <= endIndex; i++) {
            var point = _points[i];
            if (i < _tokenSpans.Length && _tokenSpans[i].HasValue) {
                var span = _tokenSpans[i]!.Value;
                start ??= new MarkdownSourcePoint(span.StartLine, span.StartColumn ?? 1, span.StartOffset ?? 0);
                end = new MarkdownSourcePoint(span.EndLine, span.EndColumn ?? 1, span.EndOffset ?? 0);
                continue;
            }

            if (point == null) {
                continue;
            }

            start ??= point;
            end = point;
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
        if (length != 1 || startIndex < 0 || startIndex >= _tokenLiterals.Length) {
            return null;
        }

        return _tokenLiterals[startIndex];
    }

    internal MarkdownInlineSourceMap Slice(int startIndex, int length) {
        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return new MarkdownInlineSourceMap(Array.Empty<MarkdownSourcePoint?>());
        }

        var actualLength = Math.Min(length, _points.Length - startIndex);
        var slice = new MarkdownSourcePoint?[actualLength];
        Array.Copy(_points, startIndex, slice, 0, actualLength);

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
