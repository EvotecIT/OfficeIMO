namespace OfficeIMO.Rtf.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private readonly Stack<HtmlFieldScope> _fieldScopes = new Stack<HtmlFieldScope>();

        private bool TryStartField(HtmlToken token) {
            string? marker = GetAttribute(token, "data-officeimo-rtf-field");
            string? instruction = GetAttribute(token, "data-officeimo-rtf-field-instruction");
            if (!IsFieldMarker(marker) && string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            RtfField field = EnsureInlineParagraph().AddField(instruction ?? string.Empty);
            _fieldScopes.Push(new HtmlFieldScope(field));
            return true;
        }

        private void EnterFieldElement() {
            if (_fieldScopes.Count > 0) {
                _fieldScopes.Peek().Depth++;
            }
        }

        private void ExitFieldElement() {
            if (_fieldScopes.Count == 0) {
                return;
            }

            HtmlFieldScope scope = _fieldScopes.Peek();
            scope.Depth--;
            if (scope.Depth <= 0) {
                _fieldScopes.Pop();
            }
        }

        private RtfParagraph EnsureInlineParagraph() {
            return _fieldScopes.Count == 0
                ? EnsureParagraph()
                : _fieldScopes.Peek().Field.Result;
        }

        private static bool IsFieldMarker(string? marker) {
            return string.Equals(marker, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(marker, "field", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(marker, "start", StringComparison.OrdinalIgnoreCase);
        }

        private sealed class HtmlFieldScope {
            internal HtmlFieldScope(RtfField field) {
                Field = field;
            }

            internal RtfField Field { get; }

            internal int Depth { get; set; } = 1;
        }
    }
}
