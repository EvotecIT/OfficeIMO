namespace OfficeIMO.Html;

internal static class HtmlCssCustomPropertyResolver {
    private const int MaximumDepth = 32;

    internal static bool TryResolve(string value, Func<string, string?> lookup, out string resolved) {
        if (lookup == null) throw new ArgumentNullException(nameof(lookup));
        return TryResolve(value ?? string.Empty, lookup, new HashSet<string>(StringComparer.OrdinalIgnoreCase), 0, out resolved);
    }

    internal static bool ContainsVarFunction(string value) =>
        !string.IsNullOrEmpty(value) && value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool TryResolve(string value, Func<string, string?> lookup, ISet<string> resolving, int depth, out string resolved) {
        resolved = value;
        if (depth > MaximumDepth) return false;
        int searchStart = 0;
        while (TryFindVarFunction(resolved, searchStart, out int start, out int open, out int close)) {
            string arguments = resolved.Substring(open + 1, close - open - 1);
            SplitArguments(arguments, out string propertyName, out string? fallback);
            if (!propertyName.StartsWith("--", StringComparison.Ordinal) || propertyName.Length <= 2) return false;

            string? replacement = null;
            bool added = resolving.Add(propertyName);
            if (added) {
                string? customValue = lookup(propertyName);
                if (customValue != null && TryResolve(customValue, lookup, resolving, depth + 1, out string customResolved)) {
                    replacement = customResolved;
                }

                resolving.Remove(propertyName);
            }

            if (replacement == null && fallback != null && TryResolve(fallback, lookup, resolving, depth + 1, out string fallbackResolved)) {
                replacement = fallbackResolved;
            }

            if (replacement == null) return false;
            resolved = resolved.Substring(0, start) + replacement + resolved.Substring(close + 1);
            searchStart = Math.Max(0, start + replacement.Length);
        }

        return true;
    }

    private static bool TryFindVarFunction(string value, int startIndex, out int start, out int open, out int close) {
        start = value.IndexOf("var(", startIndex, StringComparison.OrdinalIgnoreCase);
        if (start < 0) {
            open = close = -1;
            return false;
        }

        open = start + 3;
        close = FindMatchingParenthesis(value, open);
        return close > open;
    }

    private static int FindMatchingParenthesis(string value, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < value.Length; i++) {
            char current = value[i];
            if (quote != '\0') {
                if (current == quote && (i == 0 || value[i - 1] != '\\')) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(') {
                depth++;
            } else if (current == ')' && --depth == 0) {
                return i;
            }
        }

        return -1;
    }

    private static void SplitArguments(string arguments, out string propertyName, out string? fallback) {
        int depth = 0;
        char quote = '\0';
        for (int i = 0; i < arguments.Length; i++) {
            char current = arguments[i];
            if (quote != '\0') {
                if (current == quote && (i == 0 || arguments[i - 1] != '\\')) quote = '\0';
                continue;
            }

            if (current == '\'' || current == '"') quote = current;
            else if (current == '(') depth++;
            else if (current == ')' && depth > 0) depth--;
            else if (current == ',' && depth == 0) {
                propertyName = arguments.Substring(0, i).Trim();
                fallback = arguments.Substring(i + 1).Trim();
                return;
            }
        }

        propertyName = arguments.Trim();
        fallback = null;
    }
}
