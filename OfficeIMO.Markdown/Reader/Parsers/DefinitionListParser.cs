namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class DefinitionListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.DefinitionLists) return false;

            // Do not treat indented code blocks as definition lists.
            var first = lines[i] ?? string.Empty;
            int leading = 0; while (leading < first.Length && first[leading] == ' ') leading++;
            if (leading >= 4) return false;
            if (leading < first.Length && first[leading] == '\t') return false;

            if (!IsDefinitionLine(lines[i])) return false;
            var dl = new DefinitionListBlock();
            dl.SetParsingContext(options, state);
            int j = i;
            while (j < lines.Length && IsDefinitionLine(lines[j])) {
                var raw = lines[j] ?? string.Empty;
                int lead = 0; while (lead < raw.Length && raw[lead] == ' ') lead++;
                if (lead >= 4) break;
                if (lead < raw.Length && raw[lead] == '\t') break;

                if (!TryGetDefinitionSeparator(lines[j], out var idx)) break;
                var term = lines[j].Substring(0, idx).Trim();
                var def = lines[j].Substring(idx + 1).TrimStart();
                dl.Items.Add((term, def)); j++;
            }
            doc.Add(dl); i = j; return true;
        }
    }
}
