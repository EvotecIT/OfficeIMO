namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class CalloutParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Callouts) return false;
            if (!IsCalloutHeader(lines[i], out string kind, out string title)) return false;

            // Collect contiguous quote lines as callout body, stripping one leading ">" level.
            var inner = new System.Collections.Generic.List<string>();
            int j = i + 1;
            while (j < lines.Length) {
                var ln = lines[j] ?? string.Empty;
                var t = ln.TrimStart();
                if (!t.StartsWith(">")) break;

                if (t.Length >= 2 && t[1] == ' ') inner.Add(t.Substring(2));
                else inner.Add(t.Substring(1));
                j++;
            }

            // Parse callout body as Markdown so lists/code/etc work inside callouts.
            var nestedOptions = CloneOptionsWithoutFrontMatter(options);
            var nestedState = CloneState(state);
            var innerDoc = ParseInternal(string.Join("\n", inner), nestedOptions, nestedState, allowFrontMatter: false);
            doc.Add(new CalloutBlock(kind, title, innerDoc.Blocks));

            i = j;
            return true;
        }
    }
}
