namespace OfficeIMO.Html;

/// <summary>
/// Follows the HTML tokenizer's tag states over source text so attribute boundaries match
/// what the parser will see: quoted and unquoted values are never read as attribute names.
/// </summary>
internal static class HtmlSourceTagScanner {
    /// <summary>
    /// Scans the attributes of a start or end tag from the end of its tag name.
    /// Returns the index of the closing <c>&gt;</c>, or -1 when the input ends inside the tag.
    /// </summary>
    internal static int ScanAttributes(
        string html,
        int start,
        List<HtmlSourceAttribute> attributes,
        out bool selfClosing,
        CancellationToken cancellationToken) {
        attributes.Clear();
        selfClosing = false;
        int index = start;
        while (index < html.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            char current = html[index];
            if (IsAsciiWhitespace(current)) {
                index++;
                continue;
            }
            if (current == '>') return index;
            if (current == '/') {
                if (index + 1 < html.Length && html[index + 1] == '>') {
                    selfClosing = true;
                    return index + 1;
                }
                index++;
                continue;
            }

            // Attribute name state. A leading '=' is part of the name.
            int nameStart = index++;
            while (index < html.Length && !IsAsciiWhitespace(html[index]) && html[index] is not ('/' or '>' or '=')) index++;
            int nameEnd = index;

            // After attribute name state: only whitespace may separate the name from '='.
            int afterName = index;
            while (afterName < html.Length && IsAsciiWhitespace(html[afterName])) afterName++;
            if (afterName >= html.Length) return -1;
            if (html[afterName] != '=') {
                attributes.Add(new HtmlSourceAttribute(nameStart, nameEnd, -1, -1, nameEnd));
                index = afterName;
                continue;
            }

            // Before attribute value state.
            index = afterName + 1;
            while (index < html.Length && IsAsciiWhitespace(html[index])) index++;
            if (index >= html.Length) return -1;
            char quote = html[index];
            if (quote is '"' or '\'') {
                int valueStart = index + 1;
                int valueEnd = html.IndexOf(quote, valueStart);
                if (valueEnd < 0) return -1;
                attributes.Add(new HtmlSourceAttribute(nameStart, nameEnd, valueStart, valueEnd, valueEnd + 1));
                index = valueEnd + 1;
                continue;
            }
            if (quote == '>') {
                attributes.Add(new HtmlSourceAttribute(nameStart, nameEnd, index, index, index));
                return index;
            }

            // Unquoted value: '/' and quotes belong to the value; only whitespace or '>' end it.
            int unquotedStart = index;
            while (index < html.Length && !IsAsciiWhitespace(html[index]) && html[index] != '>') index++;
            attributes.Add(new HtmlSourceAttribute(nameStart, nameEnd, unquotedStart, index, index));
        }
        return -1;
    }

    /// <summary>Returns the first attribute with the given name, which is the one the parser keeps.</summary>
    internal static int FindAttribute(string html, List<HtmlSourceAttribute> attributes, string name) {
        for (int index = 0; index < attributes.Count; index++) {
            HtmlSourceAttribute attribute = attributes[index];
            if (attribute.NameEnd - attribute.NameStart == name.Length &&
                string.Compare(html, attribute.NameStart, name, 0, name.Length, StringComparison.OrdinalIgnoreCase) == 0) {
                return index;
            }
        }
        return -1;
    }

    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';
}

/// <summary>Source offsets of one tokenized attribute. Value offsets are -1 when the attribute has no value.</summary>
internal readonly struct HtmlSourceAttribute {
    internal HtmlSourceAttribute(int nameStart, int nameEnd, int valueStart, int valueEnd, int end) {
        NameStart = nameStart;
        NameEnd = nameEnd;
        ValueStart = valueStart;
        ValueEnd = valueEnd;
        End = end;
    }

    internal int NameStart { get; }
    internal int NameEnd { get; }
    internal int ValueStart { get; }
    internal int ValueEnd { get; }
    internal int End { get; }

    internal string Value(string html) => ValueStart < 0 ? string.Empty : html.Substring(ValueStart, ValueEnd - ValueStart);
}
