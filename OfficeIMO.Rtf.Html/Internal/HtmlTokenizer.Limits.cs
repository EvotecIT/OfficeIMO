namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlTokenizer {
    private sealed class HtmlTokenLimitTracker {
        private readonly RtfHtmlReadOptions _options;
        private readonly Stack<string> _openElements = new Stack<string>();
        private int _nodes;

        private HtmlTokenLimitTracker(RtfHtmlReadOptions options) {
            _options = options;
        }

        internal static HtmlTokenLimitTracker? Create(RtfHtmlReadOptions? options) =>
            options != null && (options.MaxHtmlNodes.HasValue || options.MaxHtmlDepth.HasValue)
                ? new HtmlTokenLimitTracker(options)
                : null;

        internal void RecordText() {
            RecordNode();
        }

        internal void RecordStart(string name, bool selfClosing) {
            RecordNode();
            int depth = _openElements.Count + 1;
            if (_options.MaxHtmlDepth.HasValue && depth > _options.MaxHtmlDepth.Value) {
                ThrowLimitExceeded("HtmlDepthLimitExceeded", "HTML nesting depth exceeded the configured conversion limit.", "MaxHtmlDepth", depth, _options.MaxHtmlDepth.Value);
            }

            if (!selfClosing) {
                _openElements.Push(name);
            }
        }

        internal void RecordEnd(string name) {
            while (_openElements.Count > 0) {
                string current = _openElements.Pop();
                if (string.Equals(current, name, StringComparison.OrdinalIgnoreCase)) {
                    break;
                }
            }
        }

        private void RecordNode() {
            _nodes++;
            if (_options.MaxHtmlNodes.HasValue && _nodes > _options.MaxHtmlNodes.Value) {
                ThrowLimitExceeded("HtmlNodeLimitExceeded", "HTML node count exceeded the configured conversion limit.", "MaxHtmlNodes", _nodes, _options.MaxHtmlNodes.Value);
            }
        }

        private void ThrowLimitExceeded(string code, string message, string source, long actual, long limit) {
            string detail = "Actual=" + actual + "; Limit=" + limit;
            var exception = new RtfHtmlConversionLimitException(code, message, source, actual, limit, detail);
            _options.AddDiagnostic(code, message, source, exception, RtfHtmlConversionDiagnosticSeverity.Error);
            throw exception;
        }
    }
}
