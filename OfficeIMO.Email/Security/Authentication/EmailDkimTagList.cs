namespace OfficeIMO.Email;

/// <summary>Strict RFC 6376 section 3.2 tag=value list parser that keeps value positions for b= removal.</summary>
internal sealed class EmailDkimTagList {
    private readonly Dictionary<string, Tag> _tags;

    private EmailDkimTagList(Dictionary<string, Tag> tags, List<string> order) {
        _tags = tags;
        Order = order;
    }

    internal IReadOnlyList<string> Order { get; }

    internal bool Contains(string name) => _tags.ContainsKey(name);

    internal string? this[string name] => _tags.TryGetValue(name, out Tag tag) ? tag.Value : null;

    internal bool TryGetSpan(string name, out int start, out int length) {
        if (_tags.TryGetValue(name, out Tag tag)) {
            start = tag.ValueStart;
            length = tag.ValueLength;
            return true;
        }

        start = 0;
        length = 0;
        return false;
    }

    /// <summary>Parses a tag list. Returns null for any syntax error, duplicate tag, or limit breach.</summary>
    internal static EmailDkimTagList? Parse(string text, int maxTags = 64) {
        var tags = new Dictionary<string, Tag>(StringComparer.Ordinal);
        var order = new List<string>();
        int segmentStart = 0;
        while (segmentStart <= text.Length) {
            int separator = text.IndexOf(';', segmentStart);
            int segmentEnd = separator < 0 ? text.Length : separator;
            bool last = separator < 0;
            if (IsFoldingWhitespaceOnly(text, segmentStart, segmentEnd)) {
                // Only one optional trailing semicolon is allowed, and the list must not be empty.
                if (!last || order.Count == 0) return null;
                break;
            }

            int equals = text.IndexOf('=', segmentStart);
            if (equals < 0 || equals >= segmentEnd) return null;
            int nameStart = SkipFws(text, segmentStart, equals);
            int nameEnd = TrimFwsEnd(text, nameStart, equals);
            string name = text.Substring(nameStart, nameEnd - nameStart);
            if (!IsTagName(name)) return null;

            int valueStart = SkipFws(text, equals + 1, segmentEnd);
            int valueEnd = TrimFwsEnd(text, valueStart, segmentEnd);
            if (!IsTagValue(text, valueStart, valueEnd)) return null;
            if (tags.ContainsKey(name) || order.Count >= maxTags) return null;
            // The span used for b= removal covers everything after '=' up to the separator, including whitespace.
            tags.Add(name, new Tag(text.Substring(valueStart, valueEnd - valueStart), equals + 1, segmentEnd - equals - 1));
            order.Add(name);
            if (last) break;
            segmentStart = segmentEnd + 1;
        }

        return order.Count == 0 ? null : new EmailDkimTagList(tags, order);
    }

    /// <summary>Removes folding whitespace from a value such as b=, bh=, or p=.</summary>
    internal static string RemoveWhitespace(string value) {
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (character != ' ' && character != '\t' && character != '\r' && character != '\n') builder.Append(character);
        }

        return builder.ToString();
    }

    internal static bool IsFws(char character) => character == ' ' || character == '\t' || character == '\r' || character == '\n';

    private static bool IsFoldingWhitespaceOnly(string text, int start, int end) {
        for (int index = start; index < end; index++) {
            if (!IsFws(text[index])) return false;
        }

        return true;
    }

    private static int SkipFws(string text, int start, int end) {
        while (start < end && IsFws(text[start])) start++;
        return start;
    }

    private static int TrimFwsEnd(string text, int start, int end) {
        while (end > start && IsFws(text[end - 1])) end--;
        return end;
    }

    private static bool IsTagName(string name) {
        if (name.Length == 0 || !IsAlpha(name[0])) return false;
        for (int index = 1; index < name.Length; index++) {
            char character = name[index];
            if (!IsAlpha(character) && !(character >= '0' && character <= '9') && character != '_') return false;
        }

        return true;
    }

    private static bool IsTagValue(string text, int start, int end) {
        for (int index = start; index < end; index++) {
            char character = text[index];
            if (character >= 0x21 && character <= 0x7E && character != ';') continue;
            if (character == ' ' || character == '\t') continue;
            // Folding: CRLF must be followed by whitespace.
            if (character == '\r' && index + 2 < end && text[index + 1] == '\n' &&
                (text[index + 2] == ' ' || text[index + 2] == '\t')) {
                index++;
                continue;
            }

            return false;
        }

        return true;
    }

    private static bool IsAlpha(char character) => (character >= 'a' && character <= 'z') || (character >= 'A' && character <= 'Z');

    private readonly struct Tag {
        internal Tag(string value, int valueStart, int valueLength) {
            Value = value;
            ValueStart = valueStart;
            ValueLength = valueLength;
        }

        internal string Value { get; }

        internal int ValueStart { get; }

        internal int ValueLength { get; }
    }
}
