namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class CalloutParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Callouts) return false;
            if (!IsCalloutHeader(lines[i], out string kind, out string title)) return false;
            var lineNumber = state.SourceLineOffset + i + 1;
            var titleSourceMap = BuildInlineSourceMapForSingleLine(
                title,
                lineNumber,
                GetCalloutTitleStartColumn(lines[i] ?? string.Empty),
                state);

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
            doc.Add(new CalloutBlock(kind, ParseInlines(title, options, state, titleSourceMap), childBlocks, syntaxChildren));

            i = j;
            return true;
        }
    }

    private static int GetCalloutTitleStartColumn(string line) {
        if (string.IsNullOrEmpty(line)) {
            return 1;
        }

        var index = 0;
        while (index < line.Length && char.IsWhiteSpace(line[index])) {
            index++;
        }

        if (index < line.Length && line[index] == '>') {
            index++;
        }

        while (index < line.Length && char.IsWhiteSpace(line[index])) {
            index++;
        }

        if (index + 1 >= line.Length || line[index] != '[' || line[index + 1] != '!') {
            return Math.Min(line.Length + 1, index + 1);
        }

        var closeIndex = line.IndexOf(']', index + 2);
        if (closeIndex < 0) {
            return Math.Min(line.Length + 1, index + 1);
        }

        index = closeIndex + 1;
        while (index < line.Length && char.IsWhiteSpace(line[index])) {
            index++;
        }

        return index + 1;
    }
}
