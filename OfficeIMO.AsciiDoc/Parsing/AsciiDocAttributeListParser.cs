namespace OfficeIMO.AsciiDoc;

internal static class AsciiDocAttributeListParser {
    internal static AsciiDocElementAttributes Parse(string content) {
        var entries = new List<AsciiDocElementAttribute>();
        int start = 0;
        char quote = '\0';
        bool escaped = false;
        for (int index = 0; index <= content.Length; index++) {
            bool atEnd = index == content.Length;
            char current = atEnd ? '\0' : content[index];
            if (!atEnd) {
                if (escaped) {
                    escaped = false;
                    continue;
                }
                if (current == '\\') {
                    escaped = true;
                    continue;
                }
                if (quote != '\0') {
                    if (current == quote) quote = '\0';
                    continue;
                }
                if (current == '\'' || current == '"') {
                    quote = current;
                    continue;
                }
            }
            if (!atEnd && current != ',') continue;

            string raw = content.Substring(start, index - start);
            string trimmed = raw.Trim();
            if (trimmed.Length > 0) entries.Add(ParseEntry(trimmed, entries.Count));
            start = index + 1;
        }
        return new AsciiDocElementAttributes(content, entries);
    }

    private static AsciiDocElementAttribute ParseEntry(string raw, int position) {
        int equals = FindUnquotedEquals(raw);
        if (equals > 0) {
            string name = raw.Substring(0, equals).Trim();
            string value = Unquote(raw.Substring(equals + 1).Trim());
            return new AsciiDocElementAttribute(AsciiDocElementAttributeKind.Named, position, raw, name, value);
        }
        if (raw.Length > 1) {
            switch (raw[0]) {
                case '#': return new AsciiDocElementAttribute(AsciiDocElementAttributeKind.Id, position, raw, null, raw.Substring(1));
                case '.': return new AsciiDocElementAttribute(AsciiDocElementAttributeKind.Role, position, raw, null, raw.Substring(1));
                case '%': return new AsciiDocElementAttribute(AsciiDocElementAttributeKind.Option, position, raw, null, raw.Substring(1));
            }
        }
        return new AsciiDocElementAttribute(AsciiDocElementAttributeKind.Positional, position, raw, null, Unquote(raw));
    }

    private static int FindUnquotedEquals(string value) {
        char quote = '\0';
        bool escaped = false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (escaped) { escaped = false; continue; }
            if (current == '\\') { escaped = true; continue; }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == '=') return index;
        }
        return -1;
    }

    private static string Unquote(string value) {
        if (value.Length >= 2 && ((value[0] == '"' && value[value.Length - 1] == '"') ||
                                  (value[0] == '\'' && value[value.Length - 1] == '\''))) {
            return value.Substring(1, value.Length - 2);
        }
        return value;
    }
}
