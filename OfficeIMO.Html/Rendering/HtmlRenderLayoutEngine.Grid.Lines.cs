using System.Globalization;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyDictionary<string, int> ParseGridLineNames(string value) {
        var names = new Dictionary<string, int>(StringComparer.Ordinal);
        int line = 0;
        AddGridLineNames(value, ref line, names);
        return names;
    }

    private static void AddGridLineNames(string value, ref int line, IDictionary<string, int> names) {
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), "none", StringComparison.OrdinalIgnoreCase)) return;
        foreach (string token in HtmlRenderCssValues.SplitWhitespace(value)) {
            string normalized = token.Trim().ToLowerInvariant();
            if (normalized.StartsWith("[", StringComparison.Ordinal) && normalized.EndsWith("]", StringComparison.Ordinal)) {
                string content = normalized.Substring(1, normalized.Length - 2);
                foreach (string name in content.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                    if (!names.ContainsKey(name)) names[name] = line;
                }
                continue;
            }

            if (normalized.StartsWith("repeat(", StringComparison.Ordinal) && normalized.EndsWith(")", StringComparison.Ordinal)) {
                IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(normalized.Substring(7, normalized.Length - 8));
                if (arguments.Count == 2
                    && int.TryParse(arguments[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out int count)
                    && count > 0) {
                    for (int iteration = 0; iteration < count; iteration++) AddGridLineNames(arguments[1], ref line, names);
                    continue;
                }
            }

            line++;
        }
    }
}
