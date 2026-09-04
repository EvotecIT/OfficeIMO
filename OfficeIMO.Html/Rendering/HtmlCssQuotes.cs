using System.Text;

namespace OfficeIMO.Html;

internal sealed class HtmlCssQuotes {
    private static readonly QuotePair[] DefaultPairs = {
        new QuotePair("\u201c", "\u201d"),
        new QuotePair("\u2018", "\u2019")
    };

    private readonly IReadOnlyList<QuotePair> _pairs;

    private HtmlCssQuotes(IReadOnlyList<QuotePair> pairs) {
        _pairs = pairs;
    }

    internal static HtmlCssQuotes Default { get; } = new HtmlCssQuotes(DefaultPairs);

    internal string OpeningAt(int depth) => _pairs.Count == 0 ? string.Empty : Resolve(depth).Opening;

    internal string ClosingAt(int depth) => _pairs.Count == 0 ? string.Empty : Resolve(depth).Closing;

    internal static bool TryParse(string? value, out HtmlCssQuotes quotes) {
        quotes = Default;
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), "auto", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string normalized = value!.Trim();
        if (string.Equals(normalized, "none", StringComparison.OrdinalIgnoreCase)) {
            quotes = new HtmlCssQuotes(Array.Empty<QuotePair>());
            return true;
        }

        var strings = new List<string>();
        int cursor = 0;
        while (cursor < normalized.Length) {
            while (cursor < normalized.Length && char.IsWhiteSpace(normalized[cursor])) cursor++;
            if (cursor >= normalized.Length) break;
            if (normalized[cursor] != '\'' && normalized[cursor] != '"') return false;
            if (!TryReadQuoted(normalized, ref cursor, out string text)) return false;
            strings.Add(text);
        }

        if (strings.Count == 0 || strings.Count % 2 != 0) return false;
        var pairs = new List<QuotePair>(strings.Count / 2);
        for (int index = 0; index < strings.Count; index += 2) {
            pairs.Add(new QuotePair(strings[index], strings[index + 1]));
        }

        quotes = new HtmlCssQuotes(pairs.AsReadOnly());
        return true;
    }

    private QuotePair Resolve(int depth) {
        int index = Math.Min(Math.Max(0, depth), _pairs.Count - 1);
        return _pairs[index];
    }

    private static bool TryReadQuoted(string value, ref int cursor, out string text) {
        char quote = value[cursor++];
        var raw = new StringBuilder();
        while (cursor < value.Length) {
            char current = value[cursor++];
            if (current == quote) {
                text = HtmlCssEscapeDecoder.Decode(raw.ToString());
                return true;
            }

            if (current == '\\' && cursor < value.Length) raw.Append(current).Append(value[cursor++]);
            else raw.Append(current);
        }

        text = string.Empty;
        return false;
    }

    private readonly struct QuotePair {
        internal QuotePair(string opening, string closing) {
            Opening = opening;
            Closing = closing;
        }

        internal string Opening { get; }
        internal string Closing { get; }
    }
}
