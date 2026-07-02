namespace OfficeIMO.Markdown;

internal static class MarkdownGenericAttributeParser {
    private static readonly IReadOnlyList<string> EmptyClasses = Array.Empty<string>();

    internal static bool TryConsumeTrailingAttributeBlock(
        string? value,
        out string textWithoutAttributeBlock,
        out MarkdownAttributeSet attributes,
        out int attributeStartIndex,
        out int attributeEndIndex,
        bool requireLeadingWhitespace = false) {
        textWithoutAttributeBlock = value ?? string.Empty;
        attributes = MarkdownAttributeSet.Empty;
        attributeStartIndex = -1;
        attributeEndIndex = -1;

        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        int end = value!.Length - 1;
        while (end >= 0 && char.IsWhiteSpace(value[end])) {
            end--;
        }

        if (end < 1 || value[end] != '}') {
            return false;
        }

        int start = FindMatchingOpeningBrace(value, end);
        if (start < 0) {
            return false;
        }

        if (requireLeadingWhitespace && start > 0 && !char.IsWhiteSpace(value[start - 1])) {
            return false;
        }

        string block = value.Substring(start + 1, end - start - 1).Trim();
        if (!TryParseAttributeBlock(block, out attributes)) {
            return false;
        }

        textWithoutAttributeBlock = value.Substring(0, start).TrimEnd();
        attributeStartIndex = start;
        attributeEndIndex = end;
        return true;
    }

    internal static bool TryConsumeLeadingAttributeBlock(
        string? value,
        out string textWithoutAttributeBlock,
        out MarkdownAttributeSet attributes,
        out int consumedLength) {
        textWithoutAttributeBlock = value ?? string.Empty;
        attributes = MarkdownAttributeSet.Empty;
        consumedLength = 0;

        if (string.IsNullOrWhiteSpace(value) || value![0] != '{') {
            return false;
        }

        int end = FindMatchingClosingBrace(value, 0);
        if (end <= 0) {
            return false;
        }

        string block = value.Substring(1, end - 1).Trim();
        if (!TryParseAttributeBlock(block, out attributes)) {
            return false;
        }

        textWithoutAttributeBlock = value.Substring(end + 1);
        consumedLength = end + 1;
        return true;
    }

    internal static bool TryParseAttributeBlock(string? value, out MarkdownAttributeSet attributes) {
        if (HasInvalidIdToken(value)) {
            attributes = MarkdownAttributeSet.Empty;
            return false;
        }

        ParseTokens(value, out var elementId, out var classes, out var attributeMap);
        attributes = MarkdownAttributeSet.Create(elementId, classes, attributeMap);
        return !attributes.IsEmpty;
    }

    internal static bool HasTrailingAttributeBlockSyntax(string? value, bool requireLeadingWhitespace = false) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        int end = value!.Length - 1;
        while (end >= 0 && char.IsWhiteSpace(value[end])) {
            end--;
        }

        if (end < 1 || value[end] != '}') {
            return false;
        }

        int start = FindMatchingOpeningBrace(value, end);
        if (start < 0) {
            return false;
        }

        return !requireLeadingWhitespace || start == 0 || char.IsWhiteSpace(value[start - 1]);
    }

    internal static void ParseTokens(
        string? value,
        out string? elementId,
        out IReadOnlyList<string> classes,
        out IReadOnlyDictionary<string, string?> attributes) {
        elementId = null;
        var parsedAttributes = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var parsedClasses = new List<string>();

        if (!string.IsNullOrWhiteSpace(value)) {
            foreach (var token in Tokenize(value!)) {
                ParseToken(token, parsedAttributes, parsedClasses, ref elementId);
            }
        }

        classes = parsedClasses.Count == 0 ? EmptyClasses : parsedClasses.AsReadOnly();
        attributes = parsedAttributes;
    }

    internal static void ParseToken(
        string token,
        IDictionary<string, string?> attributes,
        IList<string> classes,
        ref string? elementId) {
        var trimmed = token.Trim();
        if (string.IsNullOrWhiteSpace(trimmed)) {
            return;
        }

        if (trimmed[0] == '#' && trimmed.Length > 1) {
            if (string.IsNullOrWhiteSpace(elementId)) {
                elementId = trimmed.Substring(1);
            }

            return;
        }

        if (trimmed[0] == '.' && trimmed.Length > 1) {
            AddClass(classes, trimmed.Substring(1));
            return;
        }

        int equals = trimmed.IndexOf('=');
        if (equals > 0) {
            var key = trimmed.Substring(0, equals).Trim();
            if (key.Length == 0 || attributes.ContainsKey(key)) {
                return;
            }

            var rawValue = trimmed.Substring(equals + 1).Trim();
            var parsedValue = Unquote(rawValue);
            attributes[key] = parsedValue;

            if (string.Equals(key, "id", StringComparison.OrdinalIgnoreCase)
                && string.IsNullOrWhiteSpace(elementId)
                && !string.IsNullOrWhiteSpace(parsedValue)) {
                elementId = parsedValue;
                return;
            }

            if (string.Equals(key, "class", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(parsedValue)) {
                foreach (var className in parsedValue!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                    AddClass(classes, className);
                }
            }

            return;
        }

        if (!attributes.ContainsKey(trimmed)) {
            attributes[trimmed] = "true";
        }
    }

    private static bool HasInvalidIdToken(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        foreach (var token in Tokenize(value!)) {
            var trimmed = token.Trim();
            if (trimmed.Length > 0 && trimmed[0] == '#' && trimmed.Length <= 2) {
                return true;
            }
        }

        return false;
    }

    private static int FindMatchingOpeningBrace(string value, int closingBraceIndex) {
        char quote = '\0';
        int depth = 0;

        for (int i = closingBraceIndex; i >= 0; i--) {
            char ch = value[i];
            if (quote == '\0') {
                if (ch == '"' || ch == '\'') {
                    quote = ch;
                    continue;
                }

                if (ch == '}') {
                    depth++;
                    continue;
                }

                if (ch == '{') {
                    depth--;
                    if (depth == 0) {
                        return i;
                    }
                }

                continue;
            }

            if (ch == quote && (i == 0 || value[i - 1] != '\\')) {
                quote = '\0';
            }
        }

        return -1;
    }

    private static int FindMatchingClosingBrace(string value, int openingBraceIndex) {
        char quote = '\0';
        int depth = 0;

        for (int i = openingBraceIndex; i < value.Length; i++) {
            char ch = value[i];
            if (quote == '\0') {
                if (ch == '"' || ch == '\'') {
                    quote = ch;
                    continue;
                }

                if (ch == '{') {
                    depth++;
                    continue;
                }

                if (ch == '}') {
                    depth--;
                    if (depth == 0) {
                        return i;
                    }
                }

                continue;
            }

            if (ch == quote && (i == 0 || value[i - 1] != '\\')) {
                quote = '\0';
            }
        }

        return -1;
    }

    private static IEnumerable<string> Tokenize(string value) {
        var current = new StringBuilder();
        char quote = '\0';

        for (int i = 0; i < value.Length; i++) {
            var ch = value[i];
            if (quote == '\0') {
                if (char.IsWhiteSpace(ch)) {
                    if (current.Length > 0) {
                        yield return current.ToString();
                        current.Clear();
                    }

                    continue;
                }

                if (ch == '"' || ch == '\'') {
                    quote = ch;
                }

                current.Append(ch);
                continue;
            }

            current.Append(ch);
            if (ch == quote && (i == 0 || value[i - 1] != '\\')) {
                quote = '\0';
            }
        }

        if (current.Length > 0) {
            yield return current.ToString();
        }
    }

    private static string? Unquote(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return value;
        }

        if (value!.Length >= 2) {
            var first = value[0];
            var last = value[value.Length - 1];
            if ((first == '"' && last == '"') || (first == '\'' && last == '\'')) {
                return UnescapeQuotedValue(value.Substring(1, value.Length - 2), first);
            }
        }

        return value;
    }

    private static string UnescapeQuotedValue(string value, char quote) {
        if (value.IndexOf('\\') < 0) {
            return value;
        }

        var builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            var ch = value[i];
            if (ch == '\\' && i + 1 < value.Length && (value[i + 1] == quote || value[i + 1] == '\\')) {
                builder.Append(value[i + 1]);
                i++;
                continue;
            }

            builder.Append(ch);
        }

        return builder.ToString();
    }

    private static void AddClass(IList<string> classes, string? className) {
        if (string.IsNullOrWhiteSpace(className)) {
            return;
        }

        var normalized = className!.Trim();
        for (int i = 0; i < classes.Count; i++) {
            if (string.Equals(classes[i], normalized, StringComparison.OrdinalIgnoreCase)) {
                return;
            }
        }

        classes.Add(normalized);
    }
}
