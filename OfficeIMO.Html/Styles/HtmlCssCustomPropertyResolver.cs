namespace OfficeIMO.Html;

internal static class HtmlCssCustomPropertyResolver {
    private const int MaximumDepth = 32;
    private const int MaximumSyntaxDepth = 256;
    private const int MaximumSyntaxCharacters = 262144;

    internal static bool TryResolve(string value, Func<string, string?> lookup, out string resolved) {
        if (lookup == null) throw new ArgumentNullException(nameof(lookup));
        return TryResolve(value ?? string.Empty, lookup, new HashSet<string>(StringComparer.Ordinal), 0, out resolved);
    }

    internal static bool ContainsVarFunction(string value) =>
        !string.IsNullOrEmpty(value) && value.IndexOf("var(", StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool HasValidVarFunctionSyntax(string value) {
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaximumSyntaxCharacters ||
            !ContainsVarFunction(value) || !TryBuildMatchedDelimiters(value, out int[] matchingCloses)) return false;
        bool found = false;
        char quote = '\0';
        for (int index = 0; index <= value.Length - 4; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == '\\') {
                    index++;
                } else if (current == quote) {
                    quote = '\0';
                }
                continue;
            }
            if (current == '\'' || current == '"') {
                quote = current;
                continue;
            }
            if (!IsVarFunctionAt(value, index)) continue;

            found = true;
            int open = index + 3;
            int close = matchingCloses[open];
            if (close <= open) return false;
            int nameEnd = close;
            for (int argumentIndex = open + 1; argumentIndex < close; argumentIndex++) {
                if (value[argumentIndex] == '\\' && argumentIndex + 1 < close) {
                    argumentIndex++;
                    continue;
                }
                if (value[argumentIndex] == ',') {
                    nameEnd = argumentIndex;
                    break;
                }
                if (matchingCloses[argumentIndex] > argumentIndex) {
                    argumentIndex = matchingCloses[argumentIndex];
                }
            }
            string propertyName = value.Substring(open + 1, nameEnd - open - 1).Trim();
            if (!HtmlCssIdentifierParser.TryParse(propertyName, out string identifier)
                || !identifier.StartsWith("--", StringComparison.Ordinal)
                || identifier.Length <= 2) {
                return false;
            }
        }
        return found;
    }

    private static bool IsVarFunctionAt(string value, int index) {
        if (index > 0 && IsIdentifierCharacter(value[index - 1])) return false;
        return string.Compare(value, index, "var(", 0, 4, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsIdentifierCharacter(char value) =>
        char.IsLetterOrDigit(value) || value == '_' || value == '-' || value == '\\' || value >= 0x80;

    private static bool TryBuildMatchedDelimiters(string value, out int[] matchingCloses) {
        matchingCloses = new int[value.Length];
        var delimiters = new Stack<(char Open, int Index)>();
        char quote = '\0';
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote != '\0') {
                if (current == '\\') {
                    index++;
                } else if (current == quote) {
                    quote = '\0';
                }
                continue;
            }
            if (current == '\\') {
                index++;
            } else if (current == '\'' || current == '"') {
                quote = current;
            } else if (current == '(' || current == '[' || current == '{') {
                if (delimiters.Count >= MaximumSyntaxDepth) return false;
                delimiters.Push((current, index));
            } else if (current == ')' || current == ']' || current == '}') {
                if (delimiters.Count == 0) return false;
                (char open, int openIndex) = delimiters.Pop();
                if (!IsMatchingDelimiter(open, current)) return false;
                matchingCloses[openIndex] = index;
            }
        }
        return quote == '\0' && delimiters.Count == 0;
    }

    private static bool IsMatchingDelimiter(char open, char close) =>
        open == '(' && close == ')' || open == '[' && close == ']' || open == '{' && close == '}';

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
