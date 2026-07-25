using System.Globalization;
using System.Text;

namespace OfficeIMO.Html;

internal sealed class HtmlCssGeneratedContentTemplate {
    private readonly IReadOnlyList<Segment> _segments;

    private HtmlCssGeneratedContentTemplate(IEnumerable<Segment> segments) {
        _segments = new List<Segment>(segments).AsReadOnly();
    }

    internal bool IsEmpty => _segments.Count == 0;

    internal int GetRenderedLength(int pageNumber, int pageCount, HtmlCssRunningStringPageContext? runningStrings = null) {
        int length = 0;
        foreach (Segment segment in _segments) {
            int segmentLength = segment.Counter == CounterKind.Page
                ? pageNumber.ToString(CultureInfo.InvariantCulture).Length
                : segment.Counter == CounterKind.Pages
                    ? pageCount.ToString(CultureInfo.InvariantCulture).Length
                    : segment.RunningStringName != null
                        ? (runningStrings?.Resolve(segment.RunningStringName, segment.RunningStringPosition) ?? string.Empty).Length
                        : segment.Text.Length;
            length = checked(length + segmentLength);
        }
        return length;
    }

    internal string Render(int pageNumber, int pageCount, HtmlCssRunningStringPageContext? runningStrings = null) {
        var text = new StringBuilder(GetRenderedLength(pageNumber, pageCount, runningStrings));
        foreach (Segment segment in _segments) {
            if (segment.Counter == CounterKind.Page) text.Append(pageNumber.ToString(CultureInfo.InvariantCulture));
            else if (segment.Counter == CounterKind.Pages) text.Append(pageCount.ToString(CultureInfo.InvariantCulture));
            else if (segment.RunningStringName != null) text.Append(runningStrings?.Resolve(segment.RunningStringName, segment.RunningStringPosition));
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

            if (TryReadCounter(value, ref cursor, out CounterKind counter)) {
                segments.Add(new Segment(string.Empty, counter));
                continue;
            }
            if (TryReadRunningString(value, ref cursor, out string name, out HtmlCssRunningStringPosition position)) {
                segments.Add(new Segment(name, position));
                continue;
            }
            return false;
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

    private static bool TryReadRunningString(string value, ref int cursor, out string name, out HtmlCssRunningStringPosition position) {
        name = string.Empty;
        position = HtmlCssRunningStringPosition.First;
        const string prefix = "string(";
        if (cursor + prefix.Length > value.Length
            || !string.Equals(value.Substring(cursor, prefix.Length), prefix, StringComparison.OrdinalIgnoreCase)) return false;
        int close = value.IndexOf(')', cursor + prefix.Length);
        if (close < 0) return false;
        IReadOnlyList<string> arguments = HtmlRenderCssValues.SplitTopLevelCommas(value.Substring(cursor + prefix.Length, close - cursor - prefix.Length));
        if (arguments.Count is < 1 or > 2) return false;
        name = arguments[0].Trim();
        if (!IsIdentifier(name)) return false;
        if (arguments.Count == 2) {
            string keyword = arguments[1].Trim().ToLowerInvariant();
            if (keyword == "start") position = HtmlCssRunningStringPosition.Start;
            else if (keyword == "first") position = HtmlCssRunningStringPosition.First;
            else if (keyword == "last") position = HtmlCssRunningStringPosition.Last;
            else if (keyword == "first-except") position = HtmlCssRunningStringPosition.FirstExcept;
            else return false;
        }
        cursor = close + 1;
        return true;
    }

    private static bool IsIdentifier(string value) {
        if (value.Length == 0 || !(char.IsLetter(value[0]) || value[0] == '_' || value[0] == '-')) return false;
        for (int index = 1; index < value.Length; index++) {
            char current = value[index];
            if (!(char.IsLetterOrDigit(current) || current == '_' || current == '-')) return false;
        }
        return true;
    }

    private readonly struct Segment {
        internal Segment(string text, CounterKind counter) {
            Text = text;
            Counter = counter;
            RunningStringName = null;
            RunningStringPosition = HtmlCssRunningStringPosition.First;
        }

        internal Segment(string runningStringName, HtmlCssRunningStringPosition position) {
            Text = string.Empty;
            Counter = CounterKind.None;
            RunningStringName = runningStringName;
            RunningStringPosition = position;
        }

        internal string Text { get; }
        internal CounterKind Counter { get; }
        internal string? RunningStringName { get; }
        internal HtmlCssRunningStringPosition RunningStringPosition { get; }
    }

    private enum CounterKind {
        None,
        Page,
        Pages
    }
}
