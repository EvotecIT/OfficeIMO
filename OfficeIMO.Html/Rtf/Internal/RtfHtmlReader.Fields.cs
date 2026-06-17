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

            if (!IsHyperlinkFieldAllowed(token, instruction)) {
                return false;
            }

            RtfField field = EnsureInlineParagraph().AddField(instruction ?? string.Empty);
            ReadHyperlinkFieldData(token, field);
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

        private bool IsHyperlinkFieldAllowed(IElement token, string? instruction) {
            if (string.IsNullOrWhiteSpace(instruction)) {
                return AreHyperlinkFieldTargetsAllowed(token, null);
            }

            if (!TryReadHyperlinkInstructionTargets(instruction!, out IReadOnlyList<string> instructionTargets)) {
                return false;
            }

            var field = new RtfField(instruction!);
            string? instructionTarget = instructionTargets.Count == 0
                ? field.HyperlinkField?.Target?.ToString()
                : instructionTargets[0];
            return AreHyperlinkFieldTargetsAllowed(token, instructionTarget);
        }

        private bool TryReadHyperlinkInstructionTargets(string instruction, out IReadOnlyList<string> targets) {
            targets = Array.Empty<string>();
            IReadOnlyList<string> tokens = TokenizeRtfFieldInstruction(instruction);
            if (tokens.Count == 0 || !string.Equals(tokens[0], "HYPERLINK", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            var values = new List<string>();
            for (int index = 1; index < tokens.Count; index++) {
                string token = tokens[index];
                if (token.Length == 0) {
                    continue;
                }

                if (token[0] == '\\') {
                    if (RtfHyperlinkSwitchConsumesValue(token) && index + 1 < tokens.Count) {
                        index++;
                    }

                    continue;
                }

                values.Add(token);
            }

            if (values.Count > 1) {
                _options.AddDiagnostic(
                    "RtfHtmlFieldHyperlinkRejected",
                    "RTF hyperlink field instruction contains multiple targets.",
                    "data-officeimo-rtf-field-instruction");
                return false;
            }

            targets = values;
            return true;
        }

        private static bool RtfHyperlinkSwitchConsumesValue(string token) =>
            string.Equals(token, "\\l", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "\\m", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "\\n", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "\\o", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(token, "\\t", StringComparison.OrdinalIgnoreCase);

        private static IReadOnlyList<string> TokenizeRtfFieldInstruction(string instruction) {
            var tokens = new List<string>();
            int index = 0;
            while (index < instruction.Length) {
                while (index < instruction.Length && char.IsWhiteSpace(instruction[index])) {
                    index++;
                }

                if (index >= instruction.Length) {
                    break;
                }

                if (instruction[index] == '"') {
                    index++;
                    var quoted = new System.Text.StringBuilder();
                    while (index < instruction.Length) {
                        char c = instruction[index++];
                        if (c == '"') {
                            break;
                        }

                        quoted.Append(c);
                    }

                    tokens.Add(quoted.ToString());
                    continue;
                }

                int start = index;
                while (index < instruction.Length && !char.IsWhiteSpace(instruction[index])) {
                    index++;
                }

                tokens.Add(instruction.Substring(start, index - start));
            }

            return tokens;
        }

        private bool AreHyperlinkFieldTargetsAllowed(IElement token, string? instructionTarget) {
            string? explicitTarget = GetAttribute(token, "data-officeimo-rtf-field-hyperlink");
            string? href = GetAttribute(token, "href");
            if (!IsHyperlinkFieldTargetAllowed(instructionTarget, "data-officeimo-rtf-field-instruction")) {
                return false;
            }

            if (!IsHyperlinkFieldTargetAllowed(explicitTarget, "data-officeimo-rtf-field-hyperlink")) {
                return false;
            }

            if (!IsFragmentHref(href) && !IsHyperlinkFieldTargetAllowed(href, "href")) {
                return false;
            }

            return true;
        }

        private bool IsHyperlinkFieldTargetAllowed(string? target, string source) {
            if (string.IsNullOrWhiteSpace(target)) {
                return true;
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(target, _baseUri, _options.UrlPolicy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return true;
            }

            _options.AddDiagnostic(
                "RtfHtmlFieldHyperlinkRejected",
                "RTF hyperlink field target was rejected by the configured URL policy.",
                source);
            return false;
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

        private void ReadHyperlinkFieldData(IElement token, RtfField field) {
            string? explicitTarget = GetAttribute(token, "data-officeimo-rtf-field-hyperlink");
            string? href = GetAttribute(token, "href");
            string? target = explicitTarget ?? (IsFragmentHref(href) ? null : href);
            Uri? uri = ReadUriValue(target);
            if (uri != null) {
                field.Hyperlink = uri;
            }

            string? subAddress = GetAttribute(token, "data-officeimo-rtf-field-hyperlink-sub-address") ?? ReadFragmentHref(href);
            string? screenTip = GetAttribute(token, "data-officeimo-rtf-field-hyperlink-screen-tip") ?? GetAttribute(token, "title");
            string? targetFrame = GetAttribute(token, "data-officeimo-rtf-field-hyperlink-target-frame");
            string? imageMap = GetAttribute(token, "data-officeimo-rtf-field-hyperlink-image-map");
            if (subAddress == null && screenTip == null && targetFrame == null && imageMap == null) {
                return;
            }

            RtfHyperlinkFieldInfo data = field.GetOrCreateHyperlinkField();
            data.SubAddress = subAddress ?? data.SubAddress;
            data.ScreenTip = screenTip ?? data.ScreenTip;
            data.TargetFrame = targetFrame ?? data.TargetFrame;
            data.ImageMap = imageMap ?? data.ImageMap;
        }

        private static bool IsFragmentHref(string? href) => href != null && href.StartsWith("#", StringComparison.Ordinal);

        private static string? ReadFragmentHref(string? href) {
            if (!IsFragmentHref(href) || href!.Length == 1) {
                return null;
            }

            return href.Substring(1);
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
