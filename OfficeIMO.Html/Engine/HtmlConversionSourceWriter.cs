using System.IO;
using System.Text;
using AngleSharp;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>Bounds canonical source capture while retaining the provider's HTML formatting rules.</summary>
internal static class HtmlConversionSourceWriter {
    internal static string Serialize(IHtmlDocument document, HtmlConversionLimits limits) {
        using var writer = new BoundedWriter(limits);
        document.ToHtml(writer);
        return writer.ToString();
    }

    internal static void ValidateLength(long length, HtmlConversionLimits limits) {
        if (limits.MaxInputCharacters is int maximum && length > maximum)
            throw new HtmlDomLimitException(HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
                "Owned HTML data or serialized source exceeded the configured conversion limit.",
                nameof(HtmlConversionLimits.MaxInputCharacters), length, maximum);
    }

    private sealed class BoundedWriter : TextWriter {
        private readonly HtmlConversionLimits _limits;
        private readonly StringBuilder _buffer = new StringBuilder();
        internal BoundedWriter(HtmlConversionLimits limits) { _limits = limits; }
        public override Encoding Encoding => Encoding.UTF8;
        public override void Write(char value) {
            ValidateLength((long)_buffer.Length + 1, _limits);
            _buffer.Append(value);
        }
        public override void Write(string? value) {
            ValidateLength((long)_buffer.Length + (value?.Length ?? 0), _limits);
            _buffer.Append(value);
        }
        public override void Write(char[] buffer, int index, int count) {
            ValidateLength((long)_buffer.Length + count, _limits);
            _buffer.Append(buffer, index, count);
        }
        public override string ToString() => _buffer.ToString();
    }
}
