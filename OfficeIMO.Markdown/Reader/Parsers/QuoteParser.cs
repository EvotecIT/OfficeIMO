namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class QuoteParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            var t = lines[i];
            // Exclude callouts (handled earlier): they start with "> [!"
            var trimmed = t.TrimStart();
            if (!trimmed.StartsWith(">")) return false;
            if (trimmed.StartsWith(">") && trimmed.Length > 1 && trimmed[1] == ' ' && trimmed.Length > 3 && trimmed[2] == '[' && trimmed[3] == '!') return false;

            // Collect contiguous quote lines and un-prefix one ">" level
            var inner = new System.Collections.Generic.List<string>();
            int j = i;
            bool sawQuotedLine = false;
            while (j < lines.Length) {
                var ln = lines[j];
                var ltrim = ln.TrimStart();
                if (ltrim.StartsWith(">")) {
                    // Strip one level
                    if (ltrim.Length >= 2 && ltrim[1] == ' ') inner.Add(ltrim.Substring(2)); else inner.Add(ltrim.Substring(1));
                    sawQuotedLine = true;
                    j++;
                    continue;
                }

                // Lazy continuation: allow a non-quoted line to continue a blockquote paragraph
                // until a blank line followed by a non-quoted line ends the blockquote.
                if (sawQuotedLine) {
                    if (string.IsNullOrWhiteSpace(ln)) {
                        int peek = j + 1;
                        if (peek >= lines.Length) break;
                        var nextTrim = (lines[peek] ?? string.Empty).TrimStart();
                        if (!nextTrim.StartsWith(">")) break;
                        inner.Add(string.Empty);
                        j++;
                        continue;
                    }

                    // Only continue lazily when the previous inner line looks like paragraph content.
                    if (inner.Count == 0 || !LooksLikeParagraphLine(inner[inner.Count - 1])) break;

                    inner.Add(ln);
                    j++;
                    continue;
                }

                break;
            }
            // Recursively parse inner content as a separate document
            var nestedOptions = CloneOptionsWithoutFrontMatter(options);
            var nestedState = CloneState(state);
            var innerDoc = ParseInternal(string.Join("\n", inner), nestedOptions, nestedState, allowFrontMatter: false);
            var qb = new QuoteBlock();
            foreach (var b in innerDoc.Blocks) qb.Children.Add(b);
            doc.Add(qb); i = j; return true;
        }
    }

    private static bool LooksLikeParagraphLine(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;

        var t = line.TrimStart();

        // Block starters we do not want to lazily continue after.
        if (IsAtxHeading(t, out _, out _)) return false;
        if (LooksLikeHr(t)) return false;
        if (IsCodeFenceOpen(t, out _, out _, out _)) return false;
        if (LooksLikeTableRow(t)) return false;
        if (IsUnorderedListLine(t, out _, out _, out _)) return false;
        if (IsOrderedListLine(t, out _, out _)) return false;
        if (IsDefinitionLine(t)) return false;
        if (IsCalloutHeader("> " + t, out _, out _)) return false; // callout marker is quote-prefixed in source

        return true;
    }
}
