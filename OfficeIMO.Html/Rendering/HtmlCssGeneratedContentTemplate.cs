using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal sealed class HtmlCssGeneratedContentTemplate {
    private readonly IReadOnlyList<Segment> _segments;

    private HtmlCssGeneratedContentTemplate(IEnumerable<Segment> segments) {
        _segments = new List<Segment>(segments).AsReadOnly();
    }

    internal bool IsEmpty => _segments.Count == 0;

    internal int GetRenderedLength(int pageNumber, int pageCount) {
        int length = 0;
        foreach (Segment segment in _segments) {
            int segmentLength = segment.Counter == CounterKind.Page
                ? pageNumber.ToString(CultureInfo.InvariantCulture).Length
                : segment.Counter == CounterKind.Pages
                    ? pageCount.ToString(CultureInfo.InvariantCulture).Length
                    : segment.Text.Length;
            length = checked(length + segmentLength);
        }
        return length;
    }

    internal string Render(int pageNumber, int pageCount) {
        var text = new StringBuilder(GetRenderedLength(pageNumber, pageCount));
        foreach (Segment segment in _segments) {
            if (segment.Counter == CounterKind.Page) text.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
            else if (segment.Counter == CounterKind.Pages) text.Append(pageCount.ToString(CultureInfo.InvariantCulture));
            else text.Append(segment.Text);
        }

        return text.ToString();
    }

    internal static bool TryParse(string? expression, out HtmlCssGeneratedContentTemplate template) {
        template = new HtmlCssGeneratedContentTemplate(Array.Empty<Segment>());
        if (string.IsNullOrWhiteSpace(expression)) return true;
        string value = expression!.Trim();
        if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "normal", StringComparison.OrdinalIgnoreCase)) return true;

        var segments = new List<Segment>();
        int cursor = 0;
        while (cursor < value.Length) {
            while (cursor < value.Length && char.IsWhiteSpace(value[cursor])) cursor++;
            if (cursor >= value.Length) break;
            if (value[cursor] == '\'' || value[cursor] == '"') {
                if (!TryReadQuoted(value, ref cursor, out string text)) return false;
                segments.Add(new Segment(text, CounterKind.None));
                continue;
            }

            if (!TryReadCounter(value, ref cursor, out CounterKind counter)) return false;
            segments.Add(new Segment(string.Empty, counter));
        }

        template = new HtmlCssGeneratedContentTemplate(segments);
        return true;
    }

    private static bool TryReadQuoted(string value, ref int cursor, out string text) {
        char quote = value[cursor++];
        var result = new StringBuilder();
        while (cursor < value.Length) {
            char current = value[cursor++];
            if (current == quote) {
                text = result.ToString();
                return true;
            }

            if (current == '\\' && cursor < value.Length) current = value[cursor++];
            result.Append(current);
        }

        text = string.Empty;
        return false;
    }

    private static bool TryReadCounter(string value, ref int cursor, out CounterKind counter) {
        counter = CounterKind.None;
        const string prefix = "counter(";
        if (cursor + prefix.Length > value.Length
            || !string.Equals(value.Substring(cursor, prefix.Length), prefix, StringComparison.OrdinalIgnoreCase)) return false;
        int close = value.IndexOf(')', cursor + prefix.Length);
        if (close < 0) return false;
        string name = value.Substring(cursor + prefix.Length, close - cursor - prefix.Length).Trim();
        if (string.Equals(name, "page", StringComparison.OrdinalIgnoreCase)) counter = CounterKind.Page;
        else if (string.Equals(name, "pages", StringComparison.OrdinalIgnoreCase)) counter = CounterKind.Pages;
        else return false;
        cursor = close + 1;
        return true;
    }

    private readonly struct Segment {
        internal Segment(string text, CounterKind counter) {
            Text = text;
            Counter = counter;
        }

        internal string Text { get; }
        internal CounterKind Counter { get; }
    }

    private enum CounterKind {
        None,
        Page,
        Pages
    }
}
