using System.Text;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Formats events in the yaml-test-suite <c>test.event</c> notation and normalizes expected files.</summary>
public static class YamlTestEventText {
    /// <summary>
    /// Formats one event: kind, flow marker (<c>{}</c> or <c>[]</c>), <c>&amp;anchor</c>, <c>&lt;tag&gt;</c>,
    /// then the style indicator and escaped value for scalars or <c>*name</c> for aliases.
    /// </summary>
    public static string Format(YamlTestEvent value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        var text = new StringBuilder();
        switch (value.Kind) {
            case YamlTestEventKind.StreamStart: return "+STR";
            case YamlTestEventKind.StreamEnd: return "-STR";
            case YamlTestEventKind.DocumentStart: return value.IsExplicit ? "+DOC ---" : "+DOC";
            case YamlTestEventKind.DocumentEnd: return value.IsExplicit ? "-DOC ..." : "-DOC";
            case YamlTestEventKind.MappingEnd: return "-MAP";
            case YamlTestEventKind.SequenceEnd: return "-SEQ";
            case YamlTestEventKind.Alias: return "=ALI *" + value.Value;
            case YamlTestEventKind.MappingStart:
                text.Append(value.IsFlow ? "+MAP {}" : "+MAP");
                break;
            case YamlTestEventKind.SequenceStart:
                text.Append(value.IsFlow ? "+SEQ []" : "+SEQ");
                break;
            case YamlTestEventKind.Scalar:
                text.Append("=VAL");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(value), value.Kind, "Unknown event kind.");
        }

        if (value.Anchor != null) text.Append(" &").Append(value.Anchor);
        if (value.Tag != null) text.Append(" <").Append(value.Tag).Append('>');
        if (value.Kind == YamlTestEventKind.Scalar) {
            text.Append(' ').Append(StyleIndicator(value.Style));
            AppendEscaped(text, value.Value ?? string.Empty);
        }

        return text.ToString();
    }

    /// <summary>Formats a complete event stream, one event per line.</summary>
    public static IReadOnlyList<string> Format(IEnumerable<YamlTestEvent> events) {
        if (events == null) throw new ArgumentNullException(nameof(events));
        var lines = new List<string>();
        foreach (YamlTestEvent item in events) lines.Add(Format(item));
        return lines;
    }

    /// <summary>
    /// Splits expected <c>test.event</c> text into lines. Line endings are normalized and trailing empty lines removed;
    /// spaces inside a line are significant and kept.
    /// </summary>
    public static IReadOnlyList<string> Normalize(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        var lines = new List<string>();
        int start = 0;
        for (int i = 0; i <= text.Length; i++) {
            if (i < text.Length && text[i] != '\n') continue;
            int end = i;
            if (end > start && text[end - 1] == '\r') end--;
            lines.Add(text.Substring(start, end - start));
            start = i + 1;
        }

        while (lines.Count > 0 && lines[lines.Count - 1].Length == 0) lines.RemoveAt(lines.Count - 1);
        return lines;
    }

    /// <summary>
    /// Compares two event streams by content only: kinds, scalar values, tags, anchors and alias names.
    /// Scalar style, flow or block collection style and explicit document markers are presentation and are ignored.
    /// Returns the first differing index, or -1 when the streams are equivalent.
    /// </summary>
    public static int FindContentDifference(IReadOnlyList<YamlTestEvent> expected, IReadOnlyList<YamlTestEvent> actual) {
        if (expected == null) throw new ArgumentNullException(nameof(expected));
        if (actual == null) throw new ArgumentNullException(nameof(actual));
        int count = Math.Min(expected.Count, actual.Count);
        for (int i = 0; i < count; i++) {
            YamlTestEvent left = expected[i];
            YamlTestEvent right = actual[i];
            if (left.Kind != right.Kind
                || !string.Equals(left.Value, right.Value, StringComparison.Ordinal)
                || !string.Equals(left.Tag, right.Tag, StringComparison.Ordinal)
                || !string.Equals(left.Anchor, right.Anchor, StringComparison.Ordinal)) {
                return i;
            }
        }

        return expected.Count == actual.Count ? -1 : count;
    }

    private static char StyleIndicator(YamlTestScalarStyle style) => style switch {
        YamlTestScalarStyle.Plain => ':',
        YamlTestScalarStyle.SingleQuoted => '\'',
        YamlTestScalarStyle.DoubleQuoted => '"',
        YamlTestScalarStyle.Literal => '|',
        YamlTestScalarStyle.Folded => '>',
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "Unknown scalar style.")
    };

    // The suite escapes backslash and the control characters it uses; everything else is written as is.
    private static void AppendEscaped(StringBuilder text, string value) {
        foreach (char c in value) {
            switch (c) {
                case '\\': text.Append("\\\\"); break;
                case '\0': text.Append("\\0"); break;
                case '\b': text.Append("\\b"); break;
                case '\n': text.Append("\\n"); break;
                case '\r': text.Append("\\r"); break;
                case '\t': text.Append("\\t"); break;
                default: text.Append(c); break;
            }
        }
    }
}
