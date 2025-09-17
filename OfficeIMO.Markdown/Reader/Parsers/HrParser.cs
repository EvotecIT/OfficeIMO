namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HrParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            // HR is independent of options; it's safe and tiny, but follow Paragraphs toggle for symmetry
            if (!options.Paragraphs) return false;
            if (!LooksLikeHr(lines[i])) return false;
            doc.Add(new HorizontalRuleBlock());
            i++; return true;
        }
    }

    private static bool LooksLikeHr(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var t = line.Trim();
        // Allow spaces between characters, e.g., "- - -". Only '-', '*', '_' are valid.
        char kind = '\0'; int count = 0;
        for (int k = 0; k < t.Length; k++) {
            char ch = t[k];
            if (ch == ' ' || ch == '\t') continue;
            if (ch == '-' || ch == '*' || ch == '_') {
                if (kind == '\0') kind = ch; else if (kind != ch) return false;
                count++;
            } else return false;
        }
        return count >= 3;
    }
}
