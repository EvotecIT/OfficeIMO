namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class CalloutParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Callouts) return false;
            if (!IsCalloutHeader(lines[i], out string kind, out string title)) return false;

            // Collect contiguous quote lines as callout body, stripping one leading ">" level.
            var inner = new System.Collections.Generic.List<string>();
            var innerSourceLines = new System.Collections.Generic.List<MarkdownSourceLineSlice>();
            int j = i + 1;
            while (j < lines.Length) {
                var ln = lines[j] ?? string.Empty;
                var t = ln.TrimStart();
                if (!t.StartsWith(">")) break;

                var bodyLine = t.Length >= 2 && t[1] == ' ' ? t.Substring(2) : t.Substring(1);
                inner.Add(bodyLine);
                innerSourceLines.Add(new MarkdownSourceLineSlice(
                    bodyLine,
                    state.SourceLineOffset + j + 1,
                    GetQuoteContentStartColumn(ln)));
                j++;
            }

            // Parse callout body as Markdown so lists/code/etc work inside callouts.
            var (childBlocks, syntaxChildren) = ParseNestedMarkdownBlocks(innerSourceLines, options, state);
            doc.Add(new CalloutBlock(kind, ParseInlines(title, options, state), childBlocks, syntaxChildren));

            i = j;
            return true;
        }
    }
}
