using System.Globalization;

namespace OfficeIMO.Html;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private readonly Stack<HtmlFieldScope> _fieldScopes = new Stack<HtmlFieldScope>();

        private bool TryStartField(IElement token) {
            string? marker = GetAttribute(token, "data-officeimo-rtf-field");
            string? instruction = GetAttribute(token, "data-officeimo-rtf-field-instruction");
            if (!IsFieldMarker(marker) && string.IsNullOrWhiteSpace(instruction)) {
                return false;
            }

            RtfField field = EnsureInlineParagraph().AddField(instruction ?? string.Empty);
            ReadFormFieldData(token, field);
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

        private void ReadFormFieldData(IElement token, RtfField field) {
            if (!HasFormFieldData(token)) {
                return;
            }

            RtfFormFieldData data = field.GetOrCreateFormFieldData();
            ReadFormFieldControls(token, data);
            data.Name = GetAttribute(token, "data-officeimo-rtf-form-name") ?? data.Name;
            data.DefaultText = GetAttribute(token, "data-officeimo-rtf-form-default-text") ?? data.DefaultText;
            data.Format = GetAttribute(token, "data-officeimo-rtf-form-format") ?? data.Format;
            data.HelpText = GetAttribute(token, "data-officeimo-rtf-form-help-text") ?? data.HelpText;
            data.StatusText = GetAttribute(token, "data-officeimo-rtf-form-status-text") ?? data.StatusText;
            data.EntryMacro = GetAttribute(token, "data-officeimo-rtf-form-entry-macro") ?? data.EntryMacro;
            data.ExitMacro = GetAttribute(token, "data-officeimo-rtf-form-exit-macro") ?? data.ExitMacro;
            ReadFormFieldDropDownItems(token, data);
        }

        private static bool HasFormFieldData(IElement token) {
            return IsTruthy(GetAttribute(token, "data-officeimo-rtf-form-field")) ||
                   GetAttribute(token, "data-officeimo-rtf-form-controls") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-name") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-default-text") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-format") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-help-text") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-status-text") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-entry-macro") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-exit-macro") != null ||
                   GetAttribute(token, "data-officeimo-rtf-form-dropdown-items") != null;
        }

        private static void ReadFormFieldControls(IElement token, RtfFormFieldData data) {
            string? controls = GetAttribute(token, "data-officeimo-rtf-form-controls");
            if (string.IsNullOrWhiteSpace(controls)) {
                return;
            }

            string[] tokens = controls!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string rawToken in tokens) {
                string value = rawToken.Trim();
                if (value.Length == 0) {
                    continue;
                }

                int equals = value.IndexOf('=');
                string name = equals < 0 ? value : value.Substring(0, equals).Trim();
                if (!IsValidFormFieldControlName(name)) {
                    continue;
                }

                int? parameter = null;
                bool hasParameter = equals >= 0;
                if (hasParameter) {
                    string text = value.Substring(equals + 1).Trim();
                    if (text.Length > 0 && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)) {
                        parameter = parsed;
                    }
                }

                data.AddControl(name, parameter, hasParameter);
            }
        }

        private static void ReadFormFieldDropDownItems(IElement token, RtfFormFieldData data) {
            string? items = GetAttribute(token, "data-officeimo-rtf-form-dropdown-items");
            if (string.IsNullOrWhiteSpace(items)) {
                return;
            }

            string[] tokens = items!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string tokenValue in tokens) {
                string? item = DecodeBase64(tokenValue);
                if (item != null) {
                    data.AddDropDownItem(item);
                }
            }
        }

        private static string? DecodeBase64(string value) {
            try {
                return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(value));
            } catch (FormatException) {
                return null;
            }
        }

        private static bool IsValidFormFieldControlName(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            foreach (char character in value) {
                if ((character < 'A' || character > 'Z') && (character < 'a' || character > 'z')) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsTruthy(string? value) {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
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
