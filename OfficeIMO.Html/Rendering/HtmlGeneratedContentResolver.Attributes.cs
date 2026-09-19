namespace OfficeIMO.Html;

internal static partial class HtmlGeneratedContentResolver {
    internal static bool ReferencesAttribute(string? expression, string attributeName) {
        if (string.IsNullOrWhiteSpace(expression) || string.IsNullOrWhiteSpace(attributeName)) return false;
        int cursor = 0;
        while (cursor < expression!.Length) {
            while (cursor < expression.Length && char.IsWhiteSpace(expression[cursor])) cursor++;
            if (cursor >= expression.Length) break;
            if (expression[cursor] == '\'' || expression[cursor] == '"') {
                if (!TryReadQuoted(expression, ref cursor, out _)) return false;
                continue;
            }

            int tokenStart = cursor;
            if (TryReadFunction(expression, ref cursor, out string functionName, out string arguments)) {
                if (!string.Equals(functionName, "attr", StringComparison.OrdinalIgnoreCase)) continue;
                string candidate = arguments.Trim();
                int end = 0;
                while (end < candidate.Length && !char.IsWhiteSpace(candidate[end]) && candidate[end] != ',') end++;
                string referenced = HtmlCssEscapeDecoder.Decode(candidate.Substring(0, end));
                if (string.Equals(referenced, attributeName, StringComparison.OrdinalIgnoreCase)) return true;
                continue;
            }
            cursor = tokenStart;
            if (TryReadKeyword(expression, ref cursor, out _)) continue;
            cursor = tokenStart + 1;
        }
        return false;
    }
}
