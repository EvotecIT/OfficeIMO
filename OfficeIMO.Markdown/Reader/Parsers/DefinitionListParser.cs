namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class DefinitionListParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.DefinitionLists) return false;
            if (!IsDefinitionLine(lines[i])) return false;
            var dl = new DefinitionListBlock();
            int j = i;
            while (j < lines.Length && IsDefinitionLine(lines[j])) {
                var idx = lines[j].IndexOf(':');
                var term = lines[j].Substring(0, idx).Trim();
                var def = lines[j].Substring(idx + 1).TrimStart();
                dl.Items.Add((term, def)); j++;
            }
            doc.Add(dl); i = j; return true;
        }
    }
}
