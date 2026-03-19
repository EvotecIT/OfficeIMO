namespace OfficeIMO.Markdown;

internal sealed class MarkdownInlineSourceMap {
    private readonly MarkdownSourcePoint?[] _points;

    internal MarkdownInlineSourceMap(MarkdownSourcePoint?[] points) {
        _points = points ?? Array.Empty<MarkdownSourcePoint?>();
    }

    internal int Length => _points.Length;

    internal MarkdownSourceSpan? GetSpan(int startIndex, int length) {
        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return null;
        }

        var endIndex = Math.Min(_points.Length - 1, startIndex + length - 1);
        MarkdownSourcePoint? start = null;
        MarkdownSourcePoint? end = null;

        for (var i = startIndex; i <= endIndex; i++) {
            var point = _points[i];
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

    internal MarkdownInlineSourceMap Slice(int startIndex, int length) {
        if (length <= 0 || startIndex < 0 || startIndex >= _points.Length) {
            return new MarkdownInlineSourceMap(Array.Empty<MarkdownSourcePoint?>());
        }

        var actualLength = Math.Min(length, _points.Length - startIndex);
        var slice = new MarkdownSourcePoint?[actualLength];
        Array.Copy(_points, startIndex, slice, 0, actualLength);
        return new MarkdownInlineSourceMap(slice);
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
