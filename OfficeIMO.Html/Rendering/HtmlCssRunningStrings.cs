using System.Text;
using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed class HtmlCssRunningStringAssignment {
    internal HtmlCssRunningStringAssignment(
        string name,
        string value,
        double offset,
        double? orderOffset = null) {
        Name = name;
        Value = value;
        Offset = offset;
        OrderOffset = orderOffset ?? offset;
    }

    internal string Name { get; }
    internal string Value { get; }
    /// <summary>Visual block offset used for fragmentation and page placement.</summary>
    internal double Offset { get; }
    /// <summary>Monotonic document-flow offset used to resolve first and last assignments.</summary>
    internal double OrderOffset { get; }

    internal HtmlCssRunningStringAssignment Translate(double offset) =>
        new HtmlCssRunningStringAssignment(Name, Value, Offset + offset, OrderOffset + offset);

    internal HtmlCssRunningStringAssignment Place(double offset, double orderOffset) =>
        new HtmlCssRunningStringAssignment(Name, Value, offset, orderOffset);
}

internal sealed class HtmlCssRunningStringPageContext {
    private readonly Dictionary<string, string> _start;
    private readonly Dictionary<string, string> _first = new Dictionary<string, string>(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _last = new Dictionary<string, string>(StringComparer.Ordinal);

    internal HtmlCssRunningStringPageContext(IReadOnlyDictionary<string, string> current) {
        _start = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> item in current) {
            _start[item.Key] = item.Value;
        }
    }

    internal void Assign(HtmlCssRunningStringAssignment assignment, IDictionary<string, string> current) {
        if (!_first.ContainsKey(assignment.Name)) _first[assignment.Name] = assignment.Value;
        _last[assignment.Name] = assignment.Value;
        current[assignment.Name] = assignment.Value;
    }

    internal string Resolve(string name, HtmlCssRunningStringPosition position) {
        string start = _start.TryGetValue(name, out string? startValue) ? startValue : string.Empty;
        switch (position) {
            case HtmlCssRunningStringPosition.Start:
                return start;
            case HtmlCssRunningStringPosition.Last:
                return _last.TryGetValue(name, out string? last) ? last : start;
            case HtmlCssRunningStringPosition.FirstExcept:
                return _first.ContainsKey(name) ? string.Empty : start;
            default:
                return _first.TryGetValue(name, out string? first) ? first : start;
        }
    }
}

internal enum HtmlCssRunningStringPosition {
    First,
    Start,
    Last,
    FirstExcept
}

internal static class HtmlCssRunningStringParser {
    internal static IReadOnlyList<HtmlCssRunningStringAssignment> ResolveAssignments(
        IElement element,
        string? declaration,
        int maximumValueCharacters,
        Action<long> chargeOperations,
        out bool limitExceeded) {
        limitExceeded = false;
        if (string.IsNullOrWhiteSpace(declaration)) {
            return Array.Empty<HtmlCssRunningStringAssignment>();
        }
        if (string.Equals(declaration!.Trim(), "none", StringComparison.OrdinalIgnoreCase)) {
            return Array.Empty<HtmlCssRunningStringAssignment>();
        }

        var assignments = new List<HtmlCssRunningStringAssignment>();
        foreach (string item in HtmlRenderCssValues.SplitTopLevelCommas(declaration!)) {
            int cursor = 0;
            SkipWhitespace(item, ref cursor);
            if (!HtmlCssIdentifierParser.TryRead(item, ref cursor, out string name)) continue;
            SkipWhitespace(item, ref cursor);
            if (cursor >= item.Length) continue;
            if (!TryResolveContentList(
                element,
                item,
                ref cursor,
                maximumValueCharacters,
                chargeOperations,
                out string value,
                out bool itemLimitExceeded)) {
                limitExceeded |= itemLimitExceeded;
                continue;
            }
            SkipWhitespace(item, ref cursor);
            if (cursor == item.Length) {
                chargeOperations(1L);
                assignments.Add(new HtmlCssRunningStringAssignment(name, value, 0D));
            }
        }

        return assignments.AsReadOnly();
    }

    private static bool TryResolveContentList(
        IElement element,
        string text,
        ref int cursor,
        int maximumValueCharacters,
        Action<long> chargeOperations,
        out string value,
        out bool limitExceeded) {
        var result = new StringBuilder();
        bool found = false;
        limitExceeded = false;
        while (cursor < text.Length) {
            SkipWhitespace(text, ref cursor);
            if (cursor >= text.Length) break;
            if (text[cursor] == '\'' || text[cursor] == '"') {
                if (!TryReadQuoted(
                    text,
                    ref cursor,
                    maximumValueCharacters - result.Length,
                    chargeOperations,
                    out string literal,
                    out bool literalLimitExceeded)) {
                    value = string.Empty;
                    limitExceeded |= literalLimitExceeded;
                    return false;
                }
                if (!TryAppendBounded(result, literal, maximumValueCharacters)) {
                    value = string.Empty;
                    limitExceeded = true;
                    return false;
                }
                found = true;
                continue;
            }

            if (TryReadFunction(text, ref cursor, "content", out string contentArgument)) {
                if (contentArgument.Trim().Length != 0) {
                    value = string.Empty;
                    return false;
                }
                if (!TryAppendNormalizedElementText(
                    element,
                    result,
                    maximumValueCharacters,
                    chargeOperations)) {
                    value = string.Empty;
                    limitExceeded = true;
                    return false;
                }
                found = true;
                continue;
            }

            if (TryReadFunction(text, ref cursor, "attr", out string attributeName)
                && HtmlCssIdentifierParser.TryParse(attributeName.Trim(), out string decodedAttributeName)) {
                string attributeValue = element.GetAttribute(decodedAttributeName) ?? string.Empty;
                chargeOperations(attributeValue.Length);
                if (!TryAppendBounded(
                    result,
                    attributeValue,
                    maximumValueCharacters)) {
                    value = string.Empty;
                    limitExceeded = true;
                    return false;
                }
                found = true;
                continue;
            }

            value = string.Empty;
            return false;
        }

        value = result.ToString();
        return found;
    }

    private static bool TryAppendNormalizedElementText(
        IElement element,
        StringBuilder result,
        int maximumValueCharacters,
        Action<long> chargeOperations) {
        var pending = new Stack<INode>();
        for (int index = element.ChildNodes.Length - 1; index >= 0; index--) {
            pending.Push(element.ChildNodes[index]);
        }
        bool whitespace = false;
        int pendingCharge = 0;
        while (pending.Count > 0) {
            INode node = pending.Pop();
            chargeOperations(1L);
            if (node is IText textNode) {
                foreach (char current in textNode.Data) {
                    pendingCharge++;
                    if (pendingCharge == 256) {
                        chargeOperations(pendingCharge);
                        pendingCharge = 0;
                    }
                    if (char.IsWhiteSpace(current)) {
                        whitespace = result.Length > 0;
                        continue;
                    }
                    int required = whitespace ? 2 : 1;
                    if (result.Length > maximumValueCharacters - required) {
                        if (pendingCharge > 0) chargeOperations(pendingCharge);
                        return false;
                    }
                    if (whitespace) result.Append(' ');
                    result.Append(current);
                    whitespace = false;
                }
            } else {
                for (int index = node.ChildNodes.Length - 1; index >= 0; index--) {
                    pending.Push(node.ChildNodes[index]);
                }
            }
        }
        if (pendingCharge > 0) chargeOperations(pendingCharge);
        return true;
    }

    private static bool TryAppendBounded(StringBuilder result, string value, int maximumValueCharacters) {
        if (value.Length > maximumValueCharacters - result.Length) return false;
        result.Append(value);
        return true;
    }

    private static bool TryReadFunction(string text, ref int cursor, string name, out string argument) {
        argument = string.Empty;
        int start = cursor;
        if (cursor + name.Length + 2 > text.Length
            || !string.Equals(text.Substring(cursor, name.Length), name, StringComparison.OrdinalIgnoreCase)
            || text[cursor + name.Length] != '(') return false;
        int open = cursor + name.Length;
        int close = HtmlRenderCssValues.FindMatchingParenthesis(text, open);
        if (close < 0) {
            cursor = start;
            return false;
        }
        argument = text.Substring(open + 1, close - open - 1);
        cursor = close + 1;
        return true;
    }

    private static bool TryReadQuoted(
        string text,
        ref int cursor,
        int maximumCharacters,
        Action<long> chargeOperations,
        out string value,
        out bool limitExceeded) {
        char quote = text[cursor++];
        var result = new StringBuilder(Math.Min(Math.Max(0, maximumCharacters), 256));
        limitExceeded = false;
        while (cursor < text.Length) {
            char current = text[cursor++];
            if (current == quote) {
                value = result.ToString();
                return true;
            }
            if (current == '\\') {
                int backslashIndex = cursor - 1;
                HtmlCssEscapeDecoder.TryDecodeEscape(
                    text,
                    backslashIndex,
                    out string decoded,
                    out int consumedCharacters);
                cursor = backslashIndex + consumedCharacters;
                chargeOperations(consumedCharacters);
                if (!TryAppendBounded(result, decoded, maximumCharacters)) {
                    value = string.Empty;
                    limitExceeded = true;
                    return false;
                }
                continue;
            }
            chargeOperations(1L);
            if (!TryAppendBounded(result, current.ToString(), maximumCharacters)) {
                value = string.Empty;
                limitExceeded = true;
                return false;
            }
        }
        value = string.Empty;
        return false;
    }

    private static void SkipWhitespace(string text, ref int cursor) {
        while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) cursor++;
    }

}
