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
            while (j < lines.Length) {
                var ln = lines[j];
                var ltrim = ln.TrimStart();
                if (!ltrim.StartsWith(">")) break;
                // Strip one level
                if (ltrim.Length >= 2 && ltrim[1] == ' ') inner.Add(ltrim.Substring(2)); else inner.Add(ltrim.Substring(1));
                j++;
            }
            // Recursively parse inner content as a separate document
            var innerDoc = MarkdownReader.Parse(string.Join("\n", inner), options);
            var qb = new QuoteBlock();
            foreach (var b in innerDoc.Blocks) qb.Children.Add(b);
            doc.Add(qb); i = j; return true;
        }
    }
}
