using System.IO;
using System.Text;
using AngleSharp;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Bounds canonical source capture while retaining the provider's HTML formatting rules.</summary>
internal static class HtmlConversionSourceWriter {
    internal static string Serialize(IHtmlDocument document, HtmlConversionLimits limits, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        using var writer = new BoundedWriter(limits, cancellationToken);
        document.ToHtml(writer);
        string result = writer.ToString();
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    internal static void ValidateLength(long length, HtmlConversionLimits limits) {
        if (limits.MaxInputCharacters is int maximum && length > maximum)
            throw new HtmlDomLimitException(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
                "Owned HTML data or serialized source exceeded the configured conversion limit.",
                nameof(HtmlConversionLimits.MaxInputCharacters), length, maximum);
    }

    internal sealed class BoundedWriter : TextWriter {
        private readonly HtmlConversionLimits _limits;
        private readonly CancellationToken _cancellationToken;
        private readonly StringBuilder _buffer = new StringBuilder();
        private readonly bool _trimOuterWhitespace;
        private readonly StringBuilder _pendingWhitespace = new StringBuilder();
        private long _pendingWhitespaceLength;
        internal BoundedWriter(HtmlConversionLimits limits, CancellationToken cancellationToken, bool trimOuterWhitespace = false) {
            _limits = limits; _cancellationToken = cancellationToken; _trimOuterWhitespace = trimOuterWhitespace;
        }
        public override Encoding Encoding => Encoding.UTF8;
        internal int WrittenCharacters => _buffer.Length;
        internal bool IsEmpty => _buffer.Length == 0;
        internal bool EndsWithWhitespace => _pendingWhitespaceLength > 0 ||
            (_buffer.Length > 0 && char.IsWhiteSpace(_buffer[_buffer.Length - 1]));
        internal BoundedWriter Append(string? value) { Write(value); return this; }
        internal BoundedWriter Append(char value) { Write(value); return this; }
        public override void Write(char value) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_trimOuterWhitespace) { Write(value.ToString()); return; }
            ValidateLength((long)_buffer.Length + 1, _limits);
            _buffer.Append(value);
        }
        public override void Write(string? value) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_trimOuterWhitespace) { WriteTrimmed(value); return; }
            ValidateLength((long)_buffer.Length + (value?.Length ?? 0), _limits);
            _buffer.Append(value);
        }
        public override void Write(char[] buffer, int index, int count) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_trimOuterWhitespace) { Write(new string(buffer, index, count)); return; }
            ValidateLength((long)_buffer.Length + count, _limits);
            _buffer.Append(buffer, index, count);
        }

        // Normalization discards outer whitespace. Defer a possible trailing run
        // until later content makes it part of the actual serialized value.
        private void WriteTrimmed(string? value) {
            if (string.IsNullOrEmpty(value)) return;
            int start = 0;
            if (_buffer.Length == 0)
                while (start < value!.Length && char.IsWhiteSpace(value[start])) start++;
            int end = value!.Length;
            while (end > start && char.IsWhiteSpace(value[end - 1])) end--;
            if (end > start) {
                ValidateLength((long)_buffer.Length + _pendingWhitespaceLength + end - start, _limits);
                _buffer.Append(_pendingWhitespace);
                _pendingWhitespace.Clear();
                _pendingWhitespaceLength = 0;
                _buffer.Append(value, start, end - start);
            }
            if (_buffer.Length == 0) return;
            int trailing = value.Length - end;
            _pendingWhitespaceLength = checked(_pendingWhitespaceLength + trailing);
            // An over-budget pending run either disappears at the end or fails
            // before later content is appended; it need not be materialized.
            if (_limits.MaxInputCharacters is not int maximum ||
                (long)_buffer.Length + _pendingWhitespaceLength <= maximum)
                _pendingWhitespace.Append(value, end, trailing);
        }
        public override string ToString() => _buffer.ToString();
    }
}
