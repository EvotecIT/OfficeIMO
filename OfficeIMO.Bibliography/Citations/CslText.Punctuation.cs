namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    /// <summary>Resolves collisions at a rendering boundary without rewriting punctuation within a field.</summary>
    private static void MergePunctuationBoundary(StringBuilder plain, StringBuilder html, ref string next, ref string markup) {
        if (plain.Length == 0 || next.Length == 0) return;
        char left = plain[plain.Length - 1], right = next[0];
        if (":.;!?,".IndexOf(left) < 0 || ":.;!?,".IndexOf(right) < 0) return;
        bool removeRight = left == right || right == '.' && ":.;!?".IndexOf(left) >= 0 || right == ':' && ":;!?".IndexOf(left) >= 0;
        if (removeRight) {
            next = next.Substring(1);
            int index = 0;
            while (index < markup.Length && markup[index] == '<') {
                int end = markup.IndexOf('>', index + 1);
                if (end < 0) break;
                index = end + 1;
            }
            markup = markup.Remove(index, 1);
        } else if ((right == '!' || right == '?') && (left == ':' || left == ';')) {
            plain.Length--;
            int index = html.Length - 1;
            while (index >= 0 && html[index] == '>') {
                while (index >= 0 && html[index] != '<') index--;
                index--;
            }
            html.Remove(index, 1);
        }
    }

    /// <summary>Places adjoining commas and periods according to the selected locale.</summary>
    internal CslText LocalizePunctuation(CslLocale locale, int maximumCharacters, CancellationToken token) {
        if (!locale.Option("punctuation-in-quote") || !Html.Contains("data-csl-quote-end")) return this;
        string html = CslQuotePunctuation.Apply(Html, maximumCharacters, token);
        return new CslText(HtmlPlain(html, token), html, Attempted, Rendered);
    }
}
