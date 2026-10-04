using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

// Conversion metadata stays in the native owner; adapters do not lay out glyphs again.
internal sealed class XpsTextSpan {
    internal XpsTextSpan(string text, OfficePoint topLeft, OfficePoint topRight, OfficePoint bottomRight, OfficePoint bottomLeft) {
        Text = text; TopLeft = topLeft; TopRight = topRight; BottomRight = bottomRight; BottomLeft = bottomLeft;
    }
    internal string Text { get; }
    internal OfficePoint TopLeft { get; }
    internal OfficePoint TopRight { get; }
    internal OfficePoint BottomRight { get; }
    internal OfficePoint BottomLeft { get; }
    internal XpsTextSpan Transform(OfficeTransform transform) => new(Text,
        transform.TransformPoint(TopLeft), transform.TransformPoint(TopRight),
        transform.TransformPoint(BottomRight), transform.TransformPoint(BottomLeft));
}

internal sealed partial class XpsSvgConverter {
    private readonly List<XpsTextSpan> _textSpans = new();

    // Some independent PDF extractors discard standalone whitespace ActualText.
    // Attach it within the same run, before applying any native transforms.
    private void JoinWhitespace(int first, bool rtl) {
        int write = first;
        for (int read = first; read < _textSpans.Count;) {
            _token.ThrowIfCancellationRequested();
            var span = _textSpans[read++];
            var text = new StringBuilder(span.Text);
            bool onlyWhitespace = string.IsNullOrWhiteSpace(span.Text);
            double left = Math.Min(span.TopLeft.X, span.TopRight.X), right = Math.Max(span.TopLeft.X, span.TopRight.X);
            double top = span.TopLeft.Y, bottom = span.BottomLeft.Y;
            while (read < _textSpans.Count && (onlyWhitespace || string.IsNullOrWhiteSpace(_textSpans[read].Text))) {
                _token.ThrowIfCancellationRequested();
                var next = _textSpans[read++];
                text.Append(next.Text);
                onlyWhitespace &= string.IsNullOrWhiteSpace(next.Text);
                left = Math.Min(left, Math.Min(next.TopLeft.X, next.TopRight.X));
                right = Math.Max(right, Math.Max(next.TopLeft.X, next.TopRight.X));
                top = Math.Min(top, next.TopLeft.Y); bottom = Math.Max(bottom, next.BottomLeft.Y);
            }
            _textSpans[write++] = new XpsTextSpan(text.ToString(),
                new OfficePoint(rtl ? right : left, top), new OfficePoint(rtl ? left : right, top),
                new OfficePoint(rtl ? left : right, bottom), new OfficePoint(rtl ? right : left, bottom));
        }
        if (write < _textSpans.Count) _textSpans.RemoveRange(write, _textSpans.Count - write);
    }

    private void TransformText(int first, string? value) {
        if (value == null || first == _textSpans.Count) return;
        if (!OfficeSvgTransformParser.TryParse(value, out var transform)) throw new InvalidDataException("Invalid text transform.");
        if (!transform.TryInvert(out _)) { _textSpans.RemoveRange(first, _textSpans.Count - first); return; }
        for (int i = first; i < _textSpans.Count; i++) {
            _token.ThrowIfCancellationRequested();
            _textSpans[i] = _textSpans[i].Transform(transform);
        }
    }
}
