using System.Globalization;

namespace OfficeIMO.Html.Rtf;

internal static partial class RtfHtmlReader {
    private sealed partial class ReadContext {
        private void ApplyUserProperties(Dictionary<string, string> values) {
            var properties = new List<RtfUserProperty>();
            for (int index = 0; ; index++) {
                string prefix = "property." + index.ToString(CultureInfo.InvariantCulture);
                string? name = ReadString(values, prefix + ".name");
                if (string.IsNullOrWhiteSpace(name)) {
                    break;
                }

                var property = new RtfUserProperty(name!, ReadInt(values, prefix + ".typeCode"), ReadString(values, prefix + ".staticValue")) {
                    LinkedValue = ReadString(values, prefix + ".linkedValue")
                };
                properties.Add(property);
            }

            if (properties.Count > 0) {
                _document.ReplaceUserProperties(properties);
            }
        }

        private void ApplyDocumentVariables(Dictionary<string, string> values) {
            var variables = new List<RtfDocumentVariable>();
            for (int index = 0; ; index++) {
                string prefix = "variable." + index.ToString(CultureInfo.InvariantCulture);
                string? name = ReadString(values, prefix + ".name");
                if (string.IsNullOrWhiteSpace(name)) {
                    break;
                }

                string value = values.TryGetValue(prefix + ".value", out string? storedValue) ? storedValue : string.Empty;
                variables.Add(new RtfDocumentVariable(name!, value));
            }

            if (variables.Count > 0) {
                _document.ReplaceDocumentVariables(variables);
            }
        }
    }
}
