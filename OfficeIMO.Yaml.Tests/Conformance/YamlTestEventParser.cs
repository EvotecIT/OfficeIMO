using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Reads <c>test.event</c> lines back into events, the inverse of <see cref="YamlTestEventText.Format(YamlTestEvent)"/>.</summary>
public static class YamlTestEventParser {
    /// <summary>Parses a complete <c>test.event</c> text.</summary>
    public static IReadOnlyList<YamlTestEvent> Parse(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var events = new List<YamlTestEvent>();
        foreach (string line in YamlTestEventText.Normalize(text)) events.Add(ParseLine(line));
        return events;
    }

    /// <summary>Parses one event line such as <c>+MAP {} &amp;a &lt;tag&gt;</c> or <c>=VAL :value</c>.</summary>
    public static YamlTestEvent ParseLine(string line) {
        if (line == null) throw new ArgumentNullException(nameof(line));
        string kind = line.Length >= 4 ? line.Substring(0, 4) : line;
        string rest = line.Length > 5 ? line.Substring(5) : string.Empty;
        switch (kind) {
            case "+STR": return YamlTestEvent.StreamStart();
            case "-STR": return YamlTestEvent.StreamEnd();
            case "+DOC": return YamlTestEvent.DocumentStart(rest == "---");
            case "-DOC": return YamlTestEvent.DocumentEnd(rest == "...");
            case "-MAP": return YamlTestEvent.MappingEnd();
            case "-SEQ": return YamlTestEvent.SequenceEnd();
            case "=ALI":
                if (!rest.StartsWith("*", StringComparison.Ordinal)) throw Invalid(line);
                return YamlTestEvent.Alias(rest.Substring(1));
        }

        int position = 0;
        bool flow = false;
        if (rest.StartsWith("{}", StringComparison.Ordinal) || rest.StartsWith("[]", StringComparison.Ordinal)) {
            flow = true;
            position = Math.Min(rest.Length, 3);
        }

        string? anchor = null;
        if (position < rest.Length && rest[position] == '&') {
            int end = rest.IndexOf(' ', position);
            if (end < 0) end = rest.Length;
            anchor = rest.Substring(position + 1, end - position - 1);
            position = Math.Min(rest.Length, end + 1);
        }

        string? tag = null;
        if (position < rest.Length && rest[position] == '<') {
            int end = rest.IndexOf("> ", position, StringComparison.Ordinal);
            if (end < 0) end = rest.EndsWith(">", StringComparison.Ordinal) ? rest.Length - 1 : throw Invalid(line);
            tag = rest.Substring(position + 1, end - position - 1);
            position = Math.Min(rest.Length, end + 2);
        }

        switch (kind) {
            case "+MAP": return YamlTestEvent.MappingStart(flow, anchor, tag);
            case "+SEQ": return YamlTestEvent.SequenceStart(flow, anchor, tag);
            case "=VAL":
                if (position >= rest.Length) throw Invalid(line);
                YamlTestScalarStyle style = rest[position] switch {
                    ':' => YamlTestScalarStyle.Plain,
                    '\'' => YamlTestScalarStyle.SingleQuoted,
                    '"' => YamlTestScalarStyle.DoubleQuoted,
                    '|' => YamlTestScalarStyle.Literal,
                    '>' => YamlTestScalarStyle.Folded,
                    _ => throw Invalid(line)
                };
                return YamlTestEvent.Scalar(Unescape(rest.Substring(position + 1)), style, anchor, tag);
            default:
                throw Invalid(line);
        }
    }

    private static string Unescape(string value) {
        if (value.IndexOf('\\') < 0) return value;
        var text = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            if (c != '\\' || i + 1 == value.Length) {
                text.Append(c);
                continue;
            }

            char next = value[++i];
            text.Append(next switch {
                '0' => '\0',
                'b' => '\b',
                'n' => '\n',
                'r' => '\r',
                't' => '\t',
                _ => next
            });
        }

        return text.ToString();
    }

    private static FormatException Invalid(string line) => new("Unrecognized test.event line '" + line + "'.");
}
