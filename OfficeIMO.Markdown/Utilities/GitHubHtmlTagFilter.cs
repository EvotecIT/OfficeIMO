using System.Text;

namespace OfficeIMO.Markdown;

internal static class GitHubHtmlTagFilter {
    private static readonly HashSet<string> s_FilteredTagNames = new(StringComparer.OrdinalIgnoreCase) {
        "title",
        "textarea",
        "style",
        "xmp",
        "iframe",
        "noembed",
        "noframes",
        "script",
        "plaintext",
    };

    public static string Apply(string html) {
        if (string.IsNullOrEmpty(html) || html.IndexOf('<') < 0) {
            return html ?? string.Empty;
        }

        StringBuilder? builder = null;
        int lastCopyStart = 0;

        for (int i = 0; i < html.Length; i++) {
            if (html[i] != '<' || !IsFilteredTagStart(html, i)) {
                continue;
            }

            builder ??= new StringBuilder(html.Length + 16);
            builder.Append(html, lastCopyStart, i - lastCopyStart);
            builder.Append("&lt;");
            lastCopyStart = i + 1;
        }

        if (builder == null) {
            return html;
        }

        builder.Append(html, lastCopyStart, html.Length - lastCopyStart);
        return builder.ToString();
    }

    private static bool IsFilteredTagStart(string html, int tagStart) {
        int position = tagStart + 1;
        if (position >= html.Length) return false;

        if (html[position] == '/') {
            position++;
            if (position >= html.Length) return false;
        }

        if (!IsAsciiLetter(html[position])) return false;

        int nameStart = position;
        position++;
        while (position < html.Length && (IsAsciiLetter(html[position]) || char.IsDigit(html[position]) || html[position] == '-')) {
            position++;
        }

        string tagName = html.Substring(nameStart, position - nameStart);
        if (!s_FilteredTagNames.Contains(tagName)) return false;

        if (position >= html.Length) return true;

        char next = html[position];
        return next == '>' || next == '/' || char.IsWhiteSpace(next);
    }

    private static bool IsAsciiLetter(char value) =>
        (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');
}
