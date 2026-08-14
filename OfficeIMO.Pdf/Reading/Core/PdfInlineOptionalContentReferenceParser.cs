namespace OfficeIMO.Pdf;

internal sealed class PdfInlineOptionalContentReferences {
    public PdfInlineOptionalContentReferences(IReadOnlyList<PdfReference> objectReferences, bool isMembershipDictionary = false, string? policy = null, string? visibilityExpression = null, bool hasInvalidPolicy = false) {
        ObjectReferences = objectReferences;
        var objectNumbers = new int[objectReferences.Count];
        for (int index = 0; index < objectReferences.Count; index++) objectNumbers[index] = objectReferences[index].ObjectNumber;
        ObjectNumbers = objectNumbers;
        IsMembershipDictionary = isMembershipDictionary;
        Policy = string.IsNullOrWhiteSpace(policy) ? null : policy;
        VisibilityExpression = string.IsNullOrWhiteSpace(visibilityExpression) ? null : visibilityExpression;
        HasInvalidPolicy = hasInvalidPolicy;
    }

    public IReadOnlyList<int> ObjectNumbers { get; }

    public IReadOnlyList<PdfReference> ObjectReferences { get; }

    public bool IsMembershipDictionary { get; }

    public string? Policy { get; }

    public bool HasInvalidPolicy { get; }

    public string? VisibilityExpression { get; }
}

internal static class PdfInlineOptionalContentReferenceParser {
    public static PdfInlineOptionalContentReferences Read(string content, ref int index) {
        int start = index;
        SkipInlineDictionary(content, ref index);
        return Parse(content, start, Math.Max(0, index - start));
    }

    public static PdfInlineOptionalContentReferences Parse(string content, int start, int length) {
        bool isMembershipDictionary = TryReadNameValue(content, start, length, "Type", out string? type) &&
            string.Equals(type, "OCMD", StringComparison.Ordinal);
        IReadOnlyList<PdfReference> objectReferences = isMembershipDictionary &&
            TryReadObjectValue(content, start, length, "OCGs", out string? groupsValue)
                ? ExtractReferences(groupsValue!)
                : Array.Empty<PdfReference>();
        string? policyValue = null;
        bool hasPolicyValue = isMembershipDictionary && TryReadObjectValue(content, start, length, "P", out policyValue);
        string? policy = hasPolicyValue && TryReadNameValue(content, start, length, "P", out string? parsedPolicy)
            ? parsedPolicy
            : null;
        bool hasInvalidPolicy = hasPolicyValue &&
            !string.Equals(policyValue?.Trim(), "null", StringComparison.Ordinal) &&
            !IsSupportedMembershipPolicy(policy);
        string? visibilityExpression = isMembershipDictionary && TryReadObjectValue(content, start, length, "VE", out string? parsedExpression)
            ? parsedExpression
            : null;
        return new PdfInlineOptionalContentReferences(objectReferences, isMembershipDictionary, policy, visibilityExpression, hasInvalidPolicy);
    }

    private static bool IsSupportedMembershipPolicy(string? policy) =>
        policy is "AnyOn" or "AllOn" or "AnyOff" or "AllOff";

    public static IReadOnlyList<int> ExtractObjectNumbers(string content, int start, int length) {
        IReadOnlyList<PdfReference> references = ExtractReferences(content, start, length);
        if (references.Count == 0) return Array.Empty<int>();
        var objectNumbers = new int[references.Count];
        for (int index = 0; index < references.Count; index++) objectNumbers[index] = references[index].ObjectNumber;
        return objectNumbers;
    }

    private static IReadOnlyList<PdfReference> ExtractReferences(string content) =>
        ExtractReferences(content, 0, content.Length);

    private static IReadOnlyList<PdfReference> ExtractReferences(string content, int start, int length) {
        if (string.IsNullOrEmpty(content) || length <= 0 || start < 0 || start >= content.Length) {
            return Array.Empty<PdfReference>();
        }

        int end = Math.Min(content.Length, start + length);
        int index = start;
        MoveInsideOuterDictionary(content, ref index, end);
        var references = new List<PdfReference>();
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
            if (!TryReadInteger(content, ref index, end, out int generation)) {
                index = afterObjectNumber;
                continue;
            }

            SkipWhitespace(content, ref index, end);
            if (index < end && content[index] == 'R' && IsTokenBoundary(content, index + 1, end)) {
                if (objectNumber > 0 && generation >= 0) references.Add(new PdfReference(objectNumber, generation));
                index++;
            }
        }

        return references.Count == 0 ? Array.Empty<PdfReference>() : references.AsReadOnly();
    }

    public static IReadOnlyList<int> ExtractObjectNumbers(string content) =>
        ExtractObjectNumbers(content, 0, content.Length);

    private static void SkipInlineDictionary(string content, ref int index) {
        SkipInlineDictionary(content, ref index, content.Length);
    }

    private static void SkipInlineDictionary(string content, ref int index, int end) {
        if (index + 1 >= end || content[index] != '<' || content[index + 1] != '<') {
            return;
        }

        index += 2;
        int depth = 1;
        while (index < end && depth > 0) {
            char ch = content[index];
            if (ch == '(') {
                SkipLiteralString(content, ref index, end);
            } else if (ch == '<' && index + 1 < end && content[index + 1] == '<') {
                depth++;
                index += 2;
            } else if (ch == '>' && index + 1 < end && content[index + 1] == '>') {
                depth--;
                index += 2;
            } else if (ch == '<') {
                SkipHexString(content, ref index, end);
            } else if (ch == '%') {
                SkipComment(content, ref index, end);
            } else {
                index++;
            }
        }
    }

    private static void SkipLiteralString(string content, ref int index, int end) {
        int depth = 1;
        bool escaped = false;
        index++;
        while (index < end && depth > 0) {
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

    private static void SkipHexString(string content, ref int index, int end) {
        index++;
        while (index < end && content[index] != '>') {
            index++;
        }

        if (index < end) {
            index++;
        }
    }

    private static void SkipWhitespace(string content, ref int index, int end) {
        while (index < end) {
            while (index < end && IsPdfWhitespace(content[index])) {
                index++;
            }

            if (index >= end || content[index] != '%') {
                return;
            }

            SkipComment(content, ref index, end);
        }
    }

    private static void SkipComment(string content, ref int index, int end) {
        while (index < end && content[index] != '\r' && content[index] != '\n') {
            index++;
        }
    }

    private static bool IsPdfWhitespace(char ch) =>
        ch == '\0' || ch == '\t' || ch == '\n' || ch == '\f' || ch == '\r' || ch == ' ';

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
        MoveInsideOuterDictionary(content, ref index, end);
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
        MoveInsideOuterDictionary(content, ref index, end);
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

    private static void MoveInsideOuterDictionary(string content, ref int index, int end) {
        SkipWhitespace(content, ref index, end);
        if (index + 1 < end && content[index] == '<' && content[index + 1] == '<') {
            index += 2;
        }
    }

    private static string ReadNameToken(string content, ref int index, int end) {
        if (index >= end || content[index] != '/') {
            return string.Empty;
        }

        index++;
        int start = index;
        while (index < end) {
            char ch = content[index];
            if (IsPdfWhitespace(ch) ||
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

        return PdfSyntax.DecodeName(content.Substring(start, index - start));
    }

    private static void SkipToken(string content, ref int index, int end) {
        char ch = content[index];
        if (ch == '%') {
            SkipComment(content, ref index, end);
            return;
        }

        if (ch == '[' || ch == ']') {
            index++;
            return;
        }

        if (ch == '(') {
            SkipLiteralString(content, ref index, end);
            return;
        }

        if (ch == '<') {
            if (index + 1 < end && content[index + 1] == '<') {
                SkipInlineDictionary(content, ref index, end);
            } else {
                SkipHexString(content, ref index, end);
            }

            return;
        }

        index++;
        while (index < end && !IsPdfWhitespace(content[index])) {
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
        if (TrySkipIndirectReference(content, ref index, end)) {
            return;
        }

        if (ch == '[') {
            SkipArray(content, ref index, end);
        } else if (ch == '<') {
            if (index + 1 < end && content[index + 1] == '<') {
                SkipInlineDictionary(content, ref index, end);
            } else {
                SkipHexString(content, ref index, end);
            }
        } else {
            SkipToken(content, ref index, end);
        }
    }

    private static bool TrySkipIndirectReference(string content, ref int index, int end) {
        int start = index;
        if (!TryReadInteger(content, ref index, end, out _)) {
            index = start;
            return false;
        }

        SkipWhitespace(content, ref index, end);
        if (!TryReadInteger(content, ref index, end, out _)) {
            index = start;
            return false;
        }

        SkipWhitespace(content, ref index, end);
        if (index >= end || content[index] != 'R' || !IsTokenBoundary(content, index + 1, end)) {
            index = start;
            return false;
        }

        index++;
        return true;
    }

    private static bool IsTokenBoundary(string content, int index, int end) {
        if (index >= end) {
            return true;
        }

        char ch = content[index];
        return IsPdfWhitespace(ch) || ch == '%' || ch == '(' || ch == ')' || ch == '<' || ch == '>' ||
            ch == '[' || ch == ']' || ch == '{' || ch == '}' || ch == '/';
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
