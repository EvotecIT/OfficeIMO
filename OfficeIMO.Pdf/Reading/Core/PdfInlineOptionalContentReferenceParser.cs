namespace OfficeIMO.Pdf;

internal sealed class PdfInlineOptionalContentReferences {
    public PdfInlineOptionalContentReferences(IReadOnlyList<int> objectNumbers, bool isMembershipDictionary = false, string? policy = null, string? visibilityExpression = null) {
        ObjectNumbers = objectNumbers;
        IsMembershipDictionary = isMembershipDictionary;
        Policy = string.IsNullOrWhiteSpace(policy) ? null : policy;
        VisibilityExpression = string.IsNullOrWhiteSpace(visibilityExpression) ? null : visibilityExpression;
    }

    public IReadOnlyList<int> ObjectNumbers { get; }

    public bool IsMembershipDictionary { get; }

    public string? Policy { get; }

    public string? VisibilityExpression { get; }
}

internal static class PdfInlineOptionalContentReferenceParser {
    public static PdfInlineOptionalContentReferences Read(string content, ref int index) {
        int start = index;
        SkipInlineDictionary(content, ref index);
        return Parse(content, start, Math.Max(0, index - start));
    }

    public static PdfInlineOptionalContentReferences Parse(string content, int start, int length) {
        IReadOnlyList<int> objectNumbers = ExtractObjectNumbers(content, start, length);
        bool isMembershipDictionary = TryReadNameValue(content, start, length, "Type", out string? type) &&
            string.Equals(type, "OCMD", StringComparison.Ordinal);
        string? policy = isMembershipDictionary && TryReadNameValue(content, start, length, "P", out string? parsedPolicy)
            ? parsedPolicy
            : null;
        string? visibilityExpression = isMembershipDictionary && TryReadObjectValue(content, start, length, "VE", out string? parsedExpression)
            ? parsedExpression
            : null;
        return new PdfInlineOptionalContentReferences(objectNumbers, isMembershipDictionary, policy, visibilityExpression);
    }

    public static IReadOnlyList<int> ExtractObjectNumbers(string content, int start, int length) {
        if (string.IsNullOrEmpty(content) || length <= 0 || start < 0 || start >= content.Length) {
            return Array.Empty<int>();
        }

        int end = Math.Min(content.Length, start + length);
        var objectNumbers = new List<int>();
        int index = start;
        while (index < end) {
            SkipWhitespace(content, ref index, end);
            if (index >= end) {
                break;
            }

            if (!TryReadInteger(content, ref index, end, out int objectNumber)) {
                SkipToken(content, ref index, end);
                continue;
            }

            int afterObjectNumber = index;
            SkipWhitespace(content, ref index, end);
            if (!TryReadInteger(content, ref index, end, out _)) {
                index = afterObjectNumber;
                continue;
            }

            SkipWhitespace(content, ref index, end);
            if (index < end && content[index] == 'R') {
                objectNumbers.Add(objectNumber);
                index++;
            }
        }

        return objectNumbers.Count == 0 ? Array.Empty<int>() : objectNumbers.AsReadOnly();
    }

    public static IReadOnlyList<int> ExtractObjectNumbers(string content) =>
        ExtractObjectNumbers(content, 0, content.Length);

    private static void SkipInlineDictionary(string content, ref int index) {
        if (index + 1 >= content.Length || content[index] != '<' || content[index + 1] != '<') {
            return;
        }

        index += 2;
        int depth = 1;
        while (index < content.Length && depth > 0) {
            char ch = content[index];
            if (ch == '(') {
                SkipLiteralString(content, ref index);
            } else if (ch == '<' && index + 1 < content.Length && content[index + 1] == '<') {
                depth++;
                index += 2;
            } else if (ch == '>' && index + 1 < content.Length && content[index + 1] == '>') {
                depth--;
                index += 2;
            } else if (ch == '<') {
                SkipHexString(content, ref index);
            } else {
                index++;
            }
        }
    }

    private static void SkipLiteralString(string content, ref int index) {
        int depth = 1;
        bool escaped = false;
        index++;
        while (index < content.Length && depth > 0) {
            char ch = content[index++];
            if (escaped) {
                escaped = false;
            } else if (ch == '\\') {
                escaped = true;
            } else if (ch == '(') {
                depth++;
            } else if (ch == ')') {
                depth--;
            }
        }
    }

    private static void SkipHexString(string content, ref int index) {
        index++;
        while (index < content.Length && content[index] != '>') {
            index++;
        }

        if (index < content.Length) {
            index++;
        }
    }

    private static void SkipWhitespace(string content, ref int index, int end) {
        while (index < end && char.IsWhiteSpace(content[index])) {
            index++;
        }
    }

    private static bool TryReadInteger(string content, ref int index, int end, out int value) {
        value = 0;
        int start = index;
        if (index < end && (content[index] == '+' || content[index] == '-')) {
            index++;
        }

        int digitStart = index;
        while (index < end && char.IsDigit(content[index])) {
            index++;
        }

        if (index == digitStart ||
#pragma warning disable CA1846 // Keep netstandard2.0-safe parsing instead of requiring span overloads.
            !int.TryParse(content.Substring(start, index - start), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value)) {
#pragma warning restore CA1846
            index = start;
            return false;
        }

        return true;
    }

    private static bool TryReadNameValue(string content, int start, int length, string key, out string? value) {
        value = null;
        if (string.IsNullOrEmpty(content) || length <= 0 || start < 0 || start >= content.Length) {
            return false;
        }

        int end = Math.Min(content.Length, start + length);
        int index = start;
        while (index < end) {
            SkipWhitespace(content, ref index, end);
            if (index >= end) {
                return false;
            }

            if (content[index] != '/') {
                SkipToken(content, ref index, end);
                continue;
            }

            string name = ReadNameToken(content, ref index, end);
            if (!string.Equals(name, key, StringComparison.Ordinal)) {
                continue;
            }

            SkipWhitespace(content, ref index, end);
            if (index >= end || content[index] != '/') {
                return false;
            }

            value = ReadNameToken(content, ref index, end);
            return !string.IsNullOrEmpty(value);
        }

        return false;
    }

    private static bool TryReadObjectValue(string content, int start, int length, string key, out string? value) {
        value = null;
        if (string.IsNullOrEmpty(content) || length <= 0 || start < 0 || start >= content.Length) {
            return false;
        }

        int end = Math.Min(content.Length, start + length);
        int index = start;
        while (index < end) {
            SkipWhitespace(content, ref index, end);
            if (index >= end) {
                return false;
            }

            if (content[index] != '/') {
                SkipToken(content, ref index, end);
                continue;
            }

            string name = ReadNameToken(content, ref index, end);
            if (!string.Equals(name, key, StringComparison.Ordinal)) {
                SkipObject(content, ref index, end);
                continue;
            }

            SkipWhitespace(content, ref index, end);
            int valueStart = index;
            SkipObject(content, ref index, end);
            if (index <= valueStart) {
                return false;
            }

            value = content.Substring(valueStart, index - valueStart);
            return true;
        }

        return false;
    }

    private static string ReadNameToken(string content, ref int index, int end) {
        if (index >= end || content[index] != '/') {
            return string.Empty;
        }

        index++;
        int start = index;
        while (index < end) {
            char ch = content[index];
            if (char.IsWhiteSpace(ch) ||
                ch == '%' ||
                ch == '/' ||
                ch == '[' ||
                ch == ']' ||
                ch == '(' ||
                ch == ')' ||
                ch == '<' ||
                ch == '>') {
                break;
            }

            index++;
        }

        return content.Substring(start, index - start);
    }

    private static void SkipToken(string content, ref int index, int end) {
        char ch = content[index];
        if (ch == '(') {
            SkipLiteralString(content, ref index);
            return;
        }

        if (ch == '<') {
            if (index + 1 < content.Length && content[index + 1] == '<') {
                SkipInlineDictionary(content, ref index);
            } else {
                SkipHexString(content, ref index);
            }

            return;
        }

        index++;
        while (index < end && !char.IsWhiteSpace(content[index])) {
            char current = content[index];
            if (current == '[' || current == ']' || current == '/' || current == '<' || current == '>' || current == '(' || current == ')' || current == '%') {
                break;
            }

            index++;
        }
    }

    private static void SkipObject(string content, ref int index, int end) {
        SkipWhitespace(content, ref index, end);
        if (index >= end) {
            return;
        }

        char ch = content[index];
        if (ch == '[') {
            SkipArray(content, ref index, end);
        } else if (ch == '<') {
            if (index + 1 < end && content[index + 1] == '<') {
                SkipInlineDictionary(content, ref index);
            } else {
                SkipHexString(content, ref index);
            }
        } else {
            SkipToken(content, ref index, end);
        }
    }

    private static void SkipArray(string content, ref int index, int end) {
        if (index >= end || content[index] != '[') {
            return;
        }

        index++;
        while (index < end) {
            SkipWhitespace(content, ref index, end);
            if (index >= end) {
                return;
            }

            if (content[index] == ']') {
                index++;
                return;
            }

            SkipObject(content, ref index, end);
        }
    }
}
