namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static string NormalizeCodeSpanContent(string inner) {
        if (inner == null) return string.Empty;

        // Normalize newlines to spaces (CommonMark-like).
        if (inner.IndexOf('\r') >= 0) inner = inner.Replace("\r\n", "\n").Replace("\r", "\n");
        if (inner.IndexOf('\n') >= 0) inner = inner.Replace("\n", " ");

        // Trim a single leading+trailing space if both exist and the content is not all spaces.
        if (inner.Length >= 2 && inner[0] == ' ' && inner[inner.Length - 1] == ' ') {
            bool anyNonSpace = false;
            for (int i = 0; i < inner.Length; i++) {
                if (inner[i] != ' ') { anyNonSpace = true; break; }
            }
            if (anyNonSpace) inner = inner.Substring(1, inner.Length - 2);
        }

        return inner;
    }
}
